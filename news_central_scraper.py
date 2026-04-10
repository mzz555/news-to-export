#!/usr/bin/env python3
"""
中国中央主流新闻网站抓取脚本

功能：
1. 抓取人民网、新华网、央视网最新新闻
2. 提取标题、发布时间、舆情焦点、领域、链接
3. 写入 SQLite（自动去重）
4. 导出 Excel
5. 使用 schedule 每天定时执行
6. 兼容静态抓取，必要时回退 Playwright 动态抓取
"""

from __future__ import annotations

import logging
import os
import re
import sqlite3
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import schedule
from bs4 import BeautifulSoup

# Playwright 为可选依赖；若未安装，将继续使用静态抓取
try:
    from playwright.sync_api import sync_playwright
except Exception:  # noqa: BLE001
    sync_playwright = None


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/123.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

DB_PATH = "news.db"
EXCEL_PATH = "news_export.xlsx"
RUN_TIME = os.getenv("RUN_TIME", "09:00")  # 每日执行时间，可通过环境变量覆盖
REQUEST_TIMEOUT = 15
MAX_ARTICLES_PER_SITE = 20


@dataclass
class NewsItem:
    title: str
    link: str
    publish_time: str
    focus: str
    domain: str
    source_site: str
    content: str


class NewsScraper:
    def __init__(self, db_path: str = DB_PATH, excel_path: str = EXCEL_PATH) -> None:
        self.db_path = db_path
        self.excel_path = excel_path
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._init_db()

    # -------------------- 数据库 --------------------
    def _init_db(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS news (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT NOT NULL,
                    link TEXT NOT NULL UNIQUE,
                    publish_time TEXT,
                    content TEXT,
                    source_site TEXT,
                    focus TEXT,
                    domain TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            conn.commit()
        logging.info("数据库初始化完成: %s", self.db_path)

    def _existing_links(self) -> Set[str]:
        with sqlite3.connect(self.db_path) as conn:
            rows = conn.execute("SELECT link FROM news").fetchall()
        return {r[0] for r in rows}

    def _save_items(self, items: List[NewsItem]) -> int:
        inserted = 0
        with sqlite3.connect(self.db_path) as conn:
            for item in items:
                try:
                    conn.execute(
                        """
                        INSERT INTO news(title, link, publish_time, content, source_site, focus, domain)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            item.title,
                            item.link,
                            item.publish_time,
                            item.content,
                            item.source_site,
                            item.focus,
                            item.domain,
                        ),
                    )
                    inserted += 1
                except sqlite3.IntegrityError:
                    # UNIQUE(link) 冲突，自动去重
                    continue
            conn.commit()
        logging.info("数据库新增 %s 条新闻", inserted)
        return inserted

    # -------------------- 抓取基础方法 --------------------
    def _fetch_static_html(self, url: str) -> Optional[str]:
        try:
            resp = self.session.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding or "utf-8"
            return resp.text
        except Exception as e:  # noqa: BLE001
            logging.warning("静态抓取失败 %s: %s", url, e)
            return None

    def _fetch_dynamic_html(self, url: str, wait_ms: int = 2500) -> Optional[str]:
        if sync_playwright is None:
            logging.warning("Playwright 不可用，跳过动态抓取: %s", url)
            return None

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page(user_agent=USER_AGENT)
                page.goto(url, wait_until="networkidle", timeout=30000)
                page.wait_for_timeout(wait_ms)
                html = page.content()
                browser.close()
                return html
        except Exception as e:  # noqa: BLE001
            logging.warning("动态抓取失败 %s: %s", url, e)
            return None

    def _fetch_html(self, url: str) -> Optional[str]:
        html = self._fetch_static_html(url)
        if html and len(html) > 500:
            return html
        # 静态抓取不足时尝试动态
        return self._fetch_dynamic_html(url)

    @staticmethod
    def _normalize_link(base: str, href: str) -> Optional[str]:
        if not href:
            return None
        href = href.strip()
        if href.startswith("javascript:") or href.startswith("#"):
            return None
        link = urljoin(base, href)
        parsed = urlparse(link)
        if parsed.scheme not in {"http", "https"}:
            return None
        return link

    @staticmethod
    def _guess_domain_and_focus(title: str, content: str) -> Tuple[str, str]:
        text = f"{title} {content}".lower()
        rules: Dict[str, List[str]] = {
            "政治": ["国务院", "人大", "政协", "党委", "总书记", "政策", "会议"],
            "经济": ["经济", "产业", "金融", "投资", "增长", "制造业", "财政"],
            "社会": ["教育", "医疗", "养老", "就业", "民生", "社区", "治理"],
            "科技": ["科技", "人工智能", "芯片", "航天", "卫星", "机器人", "创新"],
            "国际": ["国际", "外交", "联合国", "峰会", "合作", "冲突"],
            "法治": ["法院", "检察", "公安", "法律", "执法", "反腐"],
            "文化": ["文化", "文旅", "遗产", "艺术", "博物馆", "出版"],
            "体育": ["体育", "比赛", "冠军", "奥运", "联赛"],
        }

        best_domain = "其他"
        best_score = 0
        best_words: List[str] = []

        for domain, keywords in rules.items():
            hit_words = [kw for kw in keywords if kw.lower() in text]
            if len(hit_words) > best_score:
                best_score = len(hit_words)
                best_domain = domain
                best_words = hit_words

        if best_words:
            focus = "、".join(best_words[:3])
        else:
            focus = "综合要闻"
        return best_domain, focus

    @staticmethod
    def _extract_publish_time(text: str) -> str:
        patterns = [
            r"(20\d{2}[-/.年]\d{1,2}[-/.月]\d{1,2}(?:[日\sT]\s*\d{1,2}:\d{2}(?::\d{2})?)?)",
            r"(20\d{2}-\d{2}-\d{2})",
        ]
        cleaned = re.sub(r"\s+", " ", text)
        for pat in patterns:
            m = re.search(pat, cleaned)
            if m:
                return m.group(1)
        return datetime.now().strftime("%Y-%m-%d")

    def _extract_article(self, url: str) -> Tuple[str, str]:
        html = self._fetch_html(url)
        if not html:
            return "", datetime.now().strftime("%Y-%m-%d")

        soup = BeautifulSoup(html, "lxml")
        # 优先从常见正文容器提取
        content_selectors = [
            "article",
            ".article",
            ".content",
            ".main",
            "#p-detail",
            "#content_area",
            ".detail",
            ".cnt_bd",
        ]
        content_text = ""
        for selector in content_selectors:
            node = soup.select_one(selector)
            if node:
                content_text = node.get_text(" ", strip=True)
                if len(content_text) > 100:
                    break

        if not content_text:
            content_text = soup.get_text(" ", strip=True)[:2000]

        publish_time = self._extract_publish_time(soup.get_text(" ", strip=True)[:2500])
        return content_text[:3000], publish_time

    def _extract_links_generic(
        self,
        html: str,
        base_url: str,
        source_site: str,
        allowed_domains: Iterable[str],
    ) -> List[Tuple[str, str, str]]:
        soup = BeautifulSoup(html, "lxml")
        records: List[Tuple[str, str, str]] = []
        seen: Set[str] = set()

        for a in soup.select("a[href]"):
            title = a.get_text(" ", strip=True)
            href = a.get("href", "")
            link = self._normalize_link(base_url, href)
            if not link or not title:
                continue
            if len(title) < 8:
                continue

            host = urlparse(link).netloc
            if not any(d in host for d in allowed_domains):
                continue

            # 粗筛：新闻详情页常见特征
            if not re.search(r"(\d{6,}|/\d{4}-\d{2}/\d{2}/|/\d{8}/|\.s?html?$)", link):
                continue

            key = f"{title}-{link}"
            if key in seen:
                continue
            seen.add(key)
            records.append((title, link, source_site))

            if len(records) >= MAX_ARTICLES_PER_SITE:
                break
        return records

    # -------------------- 各站点抓取 --------------------
    def scrape_people(self) -> List[NewsItem]:
        url = "http://www.people.com.cn/"
        html = self._fetch_html(url)
        if not html:
            return []

        raw = self._extract_links_generic(
            html,
            base_url=url,
            source_site="人民网",
            allowed_domains=["people.com.cn"],
        )
        return self._build_items(raw)

    def scrape_xinhua(self) -> List[NewsItem]:
        url = "http://www.news.cn/"
        html = self._fetch_html(url)
        if not html:
            return []

        raw = self._extract_links_generic(
            html,
            base_url=url,
            source_site="新华网",
            allowed_domains=["news.cn", "xinhuanet.com"],
        )
        return self._build_items(raw)

    def scrape_cctv(self) -> List[NewsItem]:
        url = "https://news.cctv.com/"
        html = self._fetch_html(url)
        if not html:
            return []

        raw = self._extract_links_generic(
            html,
            base_url=url,
            source_site="央视网",
            allowed_domains=["cctv.com"],
        )
        return self._build_items(raw)

    def _build_items(self, raw_items: List[Tuple[str, str, str]]) -> List[NewsItem]:
        items: List[NewsItem] = []
        for title, link, source in raw_items:
            try:
                content, publish_time = self._extract_article(link)
                domain, focus = self._guess_domain_and_focus(title, content)
                items.append(
                    NewsItem(
                        title=title,
                        link=link,
                        publish_time=publish_time,
                        focus=focus,
                        domain=domain,
                        source_site=source,
                        content=content,
                    )
                )
                # 减少请求压力
                time.sleep(0.2)
            except Exception as e:  # noqa: BLE001
                logging.warning("解析新闻失败 %s: %s", link, e)
                continue
        return items

    # -------------------- 导出 --------------------
    def export_to_excel(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            df = pd.read_sql_query(
                """
                SELECT id, title AS 标题, publish_time AS 发布时间, focus AS 舆情焦点,
                       domain AS 领域, link AS 链接, source_site AS 来源网站
                FROM news
                ORDER BY id DESC
                """,
                conn,
            )
        df.to_excel(self.excel_path, index=False)
        logging.info("Excel 导出完成: %s (共 %d 条)", self.excel_path, len(df))

    def run_once(self) -> None:
        logging.info("开始抓取任务: %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        existing = self._existing_links()

        all_items: List[NewsItem] = []
        for func in [self.scrape_people, self.scrape_xinhua, self.scrape_cctv]:
            try:
                site_items = func()
                # 内存级去重 + 与历史去重
                site_items = [x for x in site_items if x.link not in existing]
                for x in site_items:
                    existing.add(x.link)
                all_items.extend(site_items)
                logging.info("站点 %s 抓取到 %d 条", func.__name__, len(site_items))
            except Exception as e:  # noqa: BLE001
                logging.exception("站点抓取异常 %s: %s", func.__name__, e)

        inserted = self._save_items(all_items)
        self.export_to_excel()
        logging.info("抓取任务结束。新增 %d 条，累计候选 %d 条", inserted, len(all_items))


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("scraper.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def main() -> None:
    configure_logging()
    scraper = NewsScraper()

    # 启动时先执行一次
    try:
        scraper.run_once()
    except Exception as e:  # noqa: BLE001
        logging.exception("首次执行失败: %s", e)

    # 每日定时执行
    schedule.every().day.at(RUN_TIME).do(scraper.run_once)
    logging.info("定时任务已启动，每天 %s 执行。", RUN_TIME)

    while True:
        try:
            schedule.run_pending()
            time.sleep(1)
        except KeyboardInterrupt:
            logging.info("收到中断信号，程序退出。")
            break
        except Exception as e:  # noqa: BLE001
            logging.exception("调度循环异常: %s", e)
            time.sleep(5)


if __name__ == "__main__":
    main()
