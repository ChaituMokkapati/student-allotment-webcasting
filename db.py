"""Database helpers: MySQL (DATABASE_URL) or local SQLite fallback."""

from __future__ import annotations

import os
import sqlite3
from urllib.parse import unquote, urlparse

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "allotment.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


class _MySQLConn:
    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=None):
        sql = sql.replace("?", "%s")
        cur = self._conn.cursor()
        cur.execute(sql, params or ())
        return cur

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


def using_mysql() -> bool:
    return DATABASE_URL.startswith("mysql://") or DATABASE_URL.startswith("mysql+pymysql://")


def get_db():
    if using_mysql():
        import pymysql
        from pymysql.cursors import DictCursor

        url = DATABASE_URL.replace("mysql+pymysql://", "mysql://", 1)
        parsed = urlparse(url)
        conn = pymysql.connect(
            host=parsed.hostname or "localhost",
            port=parsed.port or 3306,
            user=unquote(parsed.username or ""),
            password=unquote(parsed.password or ""),
            database=(parsed.path or "/").lstrip("/") or "default",
            cursorclass=DictCursor,
            charset="utf8mb4",
            autocommit=False,
        )
        return _MySQLConn(conn)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def integrity_error():
    if using_mysql():
        from pymysql.err import IntegrityError

        return IntegrityError
    return sqlite3.IntegrityError


def cast_int_expr(column: str) -> str:
    if using_mysql():
        return f"CAST({column} AS UNSIGNED)"
    return f"CAST({column} AS INTEGER)"


def create_schema_sql() -> list[str]:
    if using_mysql():
        return [
            """
            CREATE TABLE IF NOT EXISTS stations (
                id INT AUTO_INCREMENT PRIMARY KEY,
                mandal VARCHAR(255) NOT NULL,
                gram_panchayat VARCHAR(255) NOT NULL,
                polling_station_no VARCHAR(64) NOT NULL,
                building_name TEXT NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS master_students (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(255) NOT NULL UNIQUE,
                year_of_study VARCHAR(64) NOT NULL,
                branch VARCHAR(64) NOT NULL,
                contact_no VARCHAR(20) NOT NULL,
                college VARCHAR(255) NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS registrations (
                id INT AUTO_INCREMENT PRIMARY KEY,
                roll_number VARCHAR(32) NOT NULL UNIQUE,
                student_name VARCHAR(255) NOT NULL,
                college VARCHAR(255) NOT NULL,
                year_of_study VARCHAR(64) NOT NULL,
                branch VARCHAR(64) NOT NULL,
                contact_no VARCHAR(20) NOT NULL,
                original_contact_no VARCHAR(20) NOT NULL,
                contact_changed TINYINT NOT NULL DEFAULT 0,
                mandal VARCHAR(255) NOT NULL,
                gram_panchayat VARCHAR(255) NOT NULL,
                polling_station_no VARCHAR(64) NOT NULL,
                building_name TEXT NOT NULL,
                created_at VARCHAR(64) NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(64) NOT NULL UNIQUE,
                password_hash VARCHAR(255) NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
        ]

    return [
        """
        CREATE TABLE IF NOT EXISTS stations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mandal TEXT NOT NULL,
            gram_panchayat TEXT NOT NULL,
            polling_station_no TEXT NOT NULL,
            building_name TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS master_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            year_of_study TEXT NOT NULL,
            branch TEXT NOT NULL,
            contact_no TEXT NOT NULL,
            college TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_number TEXT NOT NULL UNIQUE,
            student_name TEXT NOT NULL,
            college TEXT NOT NULL,
            year_of_study TEXT NOT NULL,
            branch TEXT NOT NULL,
            contact_no TEXT NOT NULL,
            original_contact_no TEXT NOT NULL,
            contact_changed INTEGER NOT NULL DEFAULT 0,
            mandal TEXT NOT NULL,
            gram_panchayat TEXT NOT NULL,
            polling_station_no TEXT NOT NULL,
            building_name TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL
        )
        """,
    ]
