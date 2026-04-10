# 中央新闻自动抓取与导出脚本

本项目提供一个可直接运行的 Python 脚本：`news_central_scraper.py`，用于每天抓取以下网站的最新新闻：

- 人民网（`people.com.cn`）
- 新华网（`news.cn`）
- 央视网（`cctv.com`）

并自动完成：

1. 抓取新闻标题、发布时间、舆情焦点、领域、链接
2. 将新闻正文和基础信息保存到 SQLite
3. 自动去重（按链接唯一）
4. 导出 Excel 表格
5. 使用 `schedule` 每天定时运行

---

## 1. 环境要求

- Python 3.9+
- 可访问目标网站网络

---

## 2. 安装依赖

### 方式 A：requirements.txt

```bash
pip install -r requirements.txt
```

### 方式 B：手动安装

```bash
pip install requests beautifulsoup4 lxml pandas openpyxl schedule playwright
```

如果要启用动态页面抓取（Playwright），还需要安装浏览器：

```bash
python -m playwright install chromium
```

> 若不安装 Playwright，脚本仍可运行，但只能做静态抓取。

---

## 3. 运行脚本

```bash
python news_central_scraper.py
```

脚本启动后会：

1. 立即执行一次抓取任务
2. 进入常驻模式，按每天固定时间执行抓取

默认每天执行时间为 `09:00`（24 小时制）。

---

## 4. 每日自动抓取配置

可以通过环境变量 `RUN_TIME` 配置每天执行时间：

### Linux / macOS

```bash
RUN_TIME=08:30 python news_central_scraper.py
```

### Windows (PowerShell)

```powershell
$env:RUN_TIME="08:30"
python news_central_scraper.py
```

---

## 5. 输出说明

### SQLite 数据库

文件：`news.db`

核心表：`news`

字段说明：

- `id`：自增主键
- `title`：新闻标题
- `link`：新闻链接（唯一，用于去重）
- `publish_time`：发布时间
- `content`：正文（截断保存）
- `source_site`：来源网站（人民网/新华网/央视网）
- `focus`：舆情焦点（关键词提取）
- `domain`：领域（政治/经济/社会/科技等）

### Excel 文件

文件：`news_export.xlsx`

列：

- `id`
- `标题`
- `发布时间`
- `舆情焦点`
- `领域`
- `链接`
- `来源网站`

---

## 6. 稳定性与异常处理

脚本已做以下处理：

- 网络请求异常捕获（超时、连接失败、HTTP 错误）
- 页面解析异常隔离（单条失败不影响整体）
- 调度循环异常保护（不中断主程序）
- 数据库唯一键冲突处理（自动跳过重复新闻）

日志输出：

- 终端实时日志
- 文件日志：`scraper.log`

---

## 7. 文件结构

```text
.
├── news_central_scraper.py   # 主脚本
├── requirements.txt          # 依赖列表
└── README.md                 # 使用说明
```

---

## 8. 注意事项

- 各网站页面结构可能变化，若抓取数量异常，请调整选择器逻辑。
- 建议控制抓取频率，避免对目标站点造成压力。
- 请遵守目标网站 robots 与相关法律法规。
