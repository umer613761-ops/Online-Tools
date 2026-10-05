import os
import re
import sqlite3
from datetime import datetime, timezone

DATABASE_URL = (os.environ.get("DATABASE_URL") or "").strip()
SQLITE_PATH = os.environ.get("WRENCHOO_RATINGS_DB", os.path.join(os.path.dirname(__file__), "wrenchoo_ratings.db"))

_PG = DATABASE_URL.startswith(("postgres://", "postgresql://", "postgresql+"))

def _pg_conn():
    import psycopg
    return psycopg.connect(DATABASE_URL)

def _sqlite_conn():
    conn = sqlite3.connect(SQLITE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_ratings_db():
    if _PG:
        with _pg_conn() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS tool_ratings (
                    id BIGSERIAL PRIMARY KEY,
                    tool_slug VARCHAR(100) NOT NULL,
                    visitor_id VARCHAR(128) NOT NULL,
                    rating INTEGER NOT NULL CHECK (rating BETWEEN 1 AND 5),
                    review TEXT NOT NULL DEFAULT '',
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    UNIQUE(tool_slug, visitor_id)
                )
            """)
            conn.commit()
    else:
        with _sqlite_conn() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS tool_ratings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tool_slug TEXT NOT NULL,
                    visitor_id TEXT NOT NULL,
                    rating INTEGER NOT NULL CHECK (rating BETWEEN 1 AND 5),
                    review TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(tool_slug, visitor_id)
                )
            """)
            conn.commit()

def save_rating(tool_slug, visitor_id, rating, review):
    review = (review or "").strip()[:500]
    if _PG:
        with _pg_conn() as conn:
            conn.execute("""
                INSERT INTO tool_ratings (tool_slug, visitor_id, rating, review)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (tool_slug, visitor_id)
                DO UPDATE SET rating = EXCLUDED.rating,
                              review = EXCLUDED.review,
                              updated_at = NOW()
            """, (tool_slug, visitor_id, rating, review))
            conn.commit()
    else:
        now = datetime.now(timezone.utc).isoformat()
        with _sqlite_conn() as conn:
            conn.execute("""
                INSERT INTO tool_ratings
                    (tool_slug, visitor_id, rating, review, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(tool_slug, visitor_id)
                DO UPDATE SET rating=excluded.rating,
                              review=excluded.review,
                              updated_at=excluded.updated_at
            """, (tool_slug, visitor_id, rating, review, now, now))
            conn.commit()

def get_rating_summary(tool_slug):
    if _PG:
        with _pg_conn() as conn:
            row = conn.execute("""
                SELECT COUNT(*) AS count, COALESCE(AVG(rating), 0) AS average
                FROM tool_ratings WHERE tool_slug = %s
            """, (tool_slug,)).fetchone()
            reviews = conn.execute("""
                SELECT rating, review, updated_at
                FROM tool_ratings
                WHERE tool_slug = %s AND review <> ''
                ORDER BY updated_at DESC LIMIT 5
            """, (tool_slug,)).fetchall()
    else:
        with _sqlite_conn() as conn:
            row = conn.execute("""
                SELECT COUNT(*) AS count, COALESCE(AVG(rating), 0) AS average
                FROM tool_ratings WHERE tool_slug = ?
            """, (tool_slug,)).fetchone()
            reviews = conn.execute("""
                SELECT rating, review, updated_at
                FROM tool_ratings
                WHERE tool_slug = ? AND review <> ''
                ORDER BY updated_at DESC LIMIT 5
            """, (tool_slug,)).fetchall()
    return {
        "count": int(row[0] if _PG else row["count"]),
        "average": round(float(row[1] if _PG else row["average"]), 1),
        "reviews": [
            {
                "rating": int(r[0]),
                "review": r[1],
                "date": str(r[2])[:10],
            } for r in reviews
        ],
    }
