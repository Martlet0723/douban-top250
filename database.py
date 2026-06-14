"""
豆瓣电影 Top250 — SQLite 数据持久化模块。

用标准库 sqlite3，零额外依赖。单文件数据库，一行 uvicorn 启动就能用。
"""

import sqlite3
from typing import Optional

DB_PATH = "movies.db"

# save_movies() 要求每条记录必须包含这些字段
_REQUIRED_FIELDS = ["排名", "电影名称", "评分"]


def _connect() -> sqlite3.Connection:
    """创建数据库连接，启用 WAL 模式 + busy_timeout 防并发锁。"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    return conn


def init_db() -> None:
    """创建 movies 表（幂等，重复调用不报错）。"""
    with _connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS movies (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                rank    INTEGER,
                title   TEXT NOT NULL,
                rating  REAL NOT NULL,
                votes   INTEGER,
                director TEXT,
                year    INTEGER,
                genre   TEXT,
                link    TEXT,
                scraped_at TEXT DEFAULT (datetime('now', 'localtime'))
            )
        """)
        conn.commit()


def clear_movies() -> None:
    """清空 movies 表（重新采集前调用）。"""
    with _connect() as conn:
        conn.execute("DELETE FROM movies")
        conn.commit()


def save_movies(movies: list[dict]) -> int:
    """
    批量插入电影数据。缺少必填字段的记录会被跳过并打印警告。

    参数:
        movies: 电影字典列表，键名与 parse_movie_item() 的返回一致。

    返回:
        实际插入的行数。
    """
    with _connect() as conn:
        count = 0
        for m in movies:
            missing = [f for f in _REQUIRED_FIELDS if m.get(f) is None]
            if missing:
                import sys
                print(f"  [警告] 跳过不完整记录: 缺少 {missing}", file=sys.stderr)
                continue

            conn.execute(
                """
                INSERT INTO movies (rank, title, rating, votes, director, year, genre, link)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    m.get("排名"),
                    m.get("电影名称"),
                    m.get("评分"),
                    m.get("评价人数"),
                    m.get("导演"),
                    m.get("年份"),
                    m.get("类型"),
                    m.get("电影链接"),
                ),
            )
            count += 1
        conn.commit()
        return count


def get_movies(min_rating: Optional[float] = None, limit: Optional[int] = 250) -> list[dict]:
    """
    查询电影列表，支持按最低评分筛选。

    参数:
        min_rating: 最低评分筛选（含），None 表示不筛选。
        limit:     最大返回条数，None 表示不限制。

    返回:
        电影字典列表。
    """
    with _connect() as conn:
        if min_rating is not None:
            sql = "SELECT * FROM movies WHERE rating >= ? ORDER BY rank"
            params = (min_rating,)
        else:
            sql = "SELECT * FROM movies ORDER BY rank"
            params = ()

        if limit is not None:
            sql += " LIMIT ?"
            params = params + (limit,)

        rows = conn.execute(sql, params)
        return [dict(row) for row in rows]


def get_stats() -> dict:
    """
    获取统计概览：总数、平均评分、最高/最低评分及对应电影。

    单条 SQL 完成，避免 N+1 查询。
    """
    with _connect() as conn:
        row = conn.execute("""
            SELECT
                COUNT(*)                       AS total,
                ROUND(AVG(rating), 2)          AS avg_rating,
                (SELECT rating FROM movies
                 ORDER BY rating DESC LIMIT 1) AS max_rating,
                (SELECT rating FROM movies
                 ORDER BY rating ASC  LIMIT 1) AS min_rating,
                (SELECT title FROM movies
                 ORDER BY rating DESC LIMIT 1) AS max_title,
                (SELECT title FROM movies
                 ORDER BY rating ASC  LIMIT 1) AS min_title
            FROM movies
        """).fetchone()

        if row is None or row["total"] == 0:
            return {"total": 0}

        return dict(row)
