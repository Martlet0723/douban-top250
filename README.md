# 豆瓣电影 Top250 数据查询系统

豆瓣电影 Top250 数据采集与查询 Web 系统。

**功能**：浏览器打开 → 点"开始采集" → 250条电影数据入库 → 表格浏览、评分筛选、导出 Excel。


## 快速开始

```bash
# 1. 安装依赖（需要 uv）
uv sync

# 2. 启动服务
uv run uvicorn app:app --reload --host 0.0.0.0 --port 8000

# 3. 打开浏览器
# http://localhost:8000
```

> 无需配置数据库，SQLite 自动创建。


## 技术栈

| 层 | 技术 |
|---|---|
| Web 框架 | FastAPI |
| 数据库 | SQLite（零配置） |
| 前端 | 原生 HTML/CSS/JS（零框架） |
| 采集 | requests + BeautifulSoup |
| 数据处理 | pandas |
| Excel 导出 | openpyxl |


## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 网页界面 |
| GET | `/api/movies?min_rating=9.0` | 电影列表（支持评分筛选） |
| GET | `/api/movies/stats` | 统计概览 |
| POST | `/api/movies/scrape` | 触发采集 |
| GET | `/api/movies/export` | 下载 Excel |

API 文档: http://localhost:8000/docs


## 项目结构

```
douban/
├── app.py              # FastAPI 入口 + 路由
├── database.py         # SQLite 操作
├── douban_top250.py    # 采集核心逻辑（复用）
├── templates/
│   └── index.html      # 前端页面
└── pyproject.toml      # 依赖管理
```
