"""
豆瓣电影 Top250 — FastAPI Web 系统入口。

启动方式:
    uv run uvicorn app:app --reload --host 0.0.0.0 --port 8000

API 文档: http://localhost:8000/docs
"""

import io
import logging
import threading
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import StreamingResponse, HTMLResponse
import pandas as pd

from database import init_db, clear_movies, save_movies, get_movies, get_stats
from douban_top250 import scrape_top250, clean_data

# ============================================================
# 初始化
# ============================================================
init_db()  # 启动时自动建表

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="豆瓣电影 Top250 数据查询系统",
    description="采集、查询、导出豆瓣电影 Top250 数据",
    version="1.0.0",
)

# ============================================================
# 读取静态首页（零依赖，不用 Jinja2）
# ============================================================
_INDEX_HTML = (Path(__file__).parent / "templates" / "index.html").read_text()

# ============================================================
# 采集进度（Web 端轮询用）
# ============================================================
_scrape_status = {
    "running": False,
    "current": 0,
    "total": 10,
    "count": 0,
}


# ============================================================
# 评分区间常量（唯一数据源，避免多处重复定义）
# ============================================================
_RATING_BINS = [0, 6.0, 7.0, 8.0, 8.5, 9.0, 9.5, 10.0]
_RATING_LABELS = ["<6.0", "6.0-6.9", "7.0-7.9", "8.0-8.4", "8.5-8.9", "9.0-9.4", "9.5-10.0"]
_FRONTEND_EXCLUDE_COLUMNS = {"id", "scraped_at"}


# ============================================================
# 工具函数：内存 Excel 导出
# ============================================================
def _movies_to_excel_bytes(movies: list[dict]) -> io.BytesIO:
    """
    将电影字典列表转为 Excel 的 BytesIO 对象（直接返回，不落文件）。

    生成两个 Sheet：
      Sheet 1 "电影数据" — 完整数据
      Sheet 2 "评分统计" — 评分区间分布 + 总览
    """
    if not movies:
        raise HTTPException(status_code=400, detail="数据库为空，请先采集数据")

    df = pd.DataFrame(movies)

    df["评分区间"] = pd.cut(
        df["rating"], bins=_RATING_BINS, labels=_RATING_LABELS, right=False
    )

    stats = (
        df.groupby("评分区间", observed=False)
        .size()
        .reset_index(name="电影数量")
    )
    stats["占比"] = (stats["电影数量"] / stats["电影数量"].sum() * 100).round(1)
    stats["占比"] = stats["占比"].astype(str) + "%"

    summary = pd.DataFrame(
        [
            ["平均评分", f"{df['rating'].mean():.2f}"],
            [
                "最高评分",
                f"{df['rating'].max():.2f}（{df.loc[df['rating'].idxmax(), 'title']}）",
            ],
            [
                "最低评分",
                f"{df['rating'].min():.2f}（{df.loc[df['rating'].idxmin(), 'title']}）",
            ],
        ],
        columns=["统计项", "数值"],
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 表头改为中文
        col_map = {
            "rank": "排名", "title": "电影名称", "rating": "评分",
            "votes": "评价人数", "director": "导演", "year": "年份",
            "genre": "类型", "link": "电影链接",
        }
        df_out = df.drop(columns=["评分区间", "id", "scraped_at"], errors="ignore").rename(columns=col_map)
        df_out.to_excel(writer, sheet_name="电影数据", index=False)
        stats.to_excel(writer, sheet_name="评分统计", index=False)
        summary.to_excel(writer, sheet_name="评分统计", startcol=4, index=False)

    output.seek(0)
    return output


# ============================================================
# 路由：首页
# ============================================================
@app.get("/", response_class=HTMLResponse)
def index():
    """返回前端页面。"""
    return HTMLResponse(_INDEX_HTML)


@app.get("/favicon.ico")
def favicon():
    """消除 404 噪音。"""
    return HTMLResponse(status_code=204)


# ============================================================
# 路由：电影列表
# ============================================================
@app.get("/api/movies")
def api_movies(min_rating: Optional[float] = Query(None, ge=0, le=10, description="最低评分筛选")):
    """
    返回电影列表（JSON 格式）。

    支持 ?min_rating=9.0 筛选评分不低于指定值的电影。
    """
    movies = get_movies(min_rating=min_rating)
    # 过滤掉内部列（id, scraped_at），只返回展示字段
    return [
        {k: v for k, v in m.items() if k not in _FRONTEND_EXCLUDE_COLUMNS}
        for m in movies
    ]


# ============================================================
# 路由：统计
# ============================================================
@app.get("/api/movies/stats")
def api_stats():
    """返回统计概览（总数、平均评分、最高/最低）。"""
    return get_stats()


# ============================================================
# 路由：触发采集
# ============================================================
@app.post("/api/movies/scrape")
def api_scrape():
    """
    触发一次新的采集任务（后台线程执行，前端轮询 /api/movies/scrape/status 获取进度）。
    """
    global _scrape_status
    if _scrape_status["running"]:
        raise HTTPException(status_code=409, detail="采集正在进行中，请等待完成")

    _scrape_status = {"running": True, "current": 0, "total": 10, "count": 0}

    def _run():
        global _scrape_status
        try:
            def on_progress(current, total, count):
                _scrape_status["current"] = current
                _scrape_status["total"] = total
                _scrape_status["count"] = count

            movies = scrape_top250(progress_callback=on_progress)
            df = clean_data(movies)
            records = df.to_dict(orient="records")
            clear_movies()
            save_movies(records)
            _scrape_status["running"] = False
            logger.info(f"采集完成: {len(records)} 条入库")
        except Exception:
            logger.exception("采集失败")
            _scrape_status["running"] = False
            _scrape_status["current"] = -1  # 标记失败

    threading.Thread(target=_run, daemon=True).start()
    return {"status": "started", "message": "采集任务已启动"}


# ============================================================
# 路由：采集进度
# ============================================================
@app.get("/api/movies/scrape/status")
def api_scrape_status():
    """返回当前采集进度（前端轮询）。"""
    return _scrape_status


# ============================================================
# 路由：导出 Excel
# ============================================================
@app.get("/api/movies/export")
def api_export():
    """
    导出 Excel 文件（.xlsx）。

    直接返回文件流，浏览器会自动触发下载。
    """
    movies = get_movies(limit=None)
    data = _movies_to_excel_bytes(movies)
    return StreamingResponse(
        data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=douban_top250.xlsx"},
    )
