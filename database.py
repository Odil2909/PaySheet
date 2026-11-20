from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


DATABASE_PATH = Path(
    os.getenv("PAYROLL_DB_PATH", str(Path(__file__).with_name("payroll.sqlite3")))
)


class ConcurrentPayrollUpdateError(RuntimeError):
    pass


def _connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 30000")
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize(config_admin_ids: set[int] | None = None) -> None:
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _connect() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY,
                normalized_name TEXT NOT NULL UNIQUE,
                display_name TEXT NOT NULL,
                password_hash BLOB,
                telegram_id INTEGER UNIQUE
            );
            CREATE TABLE IF NOT EXISTS periods (
                week_id TEXT PRIMARY KEY,
                headers_json TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS payroll_rows (
                week_id TEXT NOT NULL REFERENCES periods(week_id) ON DELETE CASCADE,
                employee_id INTEGER NOT NULL REFERENCES employees(id),
                values_json TEXT NOT NULL,
                PRIMARY KEY (week_id, employee_id)
            );
            CREATE TABLE IF NOT EXISTS login_attempts (
                telegram_id INTEGER PRIMARY KEY,
                failures INTEGER NOT NULL,
                window_started REAL NOT NULL,
                locked_until REAL NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS request_log (
                id INTEGER PRIMARY KEY,
                telegram_id INTEGER NOT NULL,
                action TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS admins (
                telegram_id INTEGER PRIMARY KEY,
                source TEXT NOT NULL CHECK (source IN ('config', 'database')),
                added_by INTEGER,
                added_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS fsm_storage (
                storage_key TEXT PRIMARY KEY,
                state TEXT,
                data_json TEXT NOT NULL DEFAULT '{}'
            );
            CREATE TABLE IF NOT EXISTS payroll_upload_staging (
                token TEXT PRIMARY KEY,
                admin_id INTEGER NOT NULL,
                week_id TEXT NOT NULL,
                headers_json TEXT NOT NULL,
                rows_json TEXT NOT NULL,
                work_rows_json TEXT NOT NULL DEFAULT '[]',
                expected_uploaded_at TEXT,
                created_at REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS employee_work_rows (
                week_id TEXT NOT NULL REFERENCES periods(week_id) ON DELETE CASCADE,
                employee_id INTEGER NOT NULL REFERENCES employees(id),
                row_order INTEGER NOT NULL,
                work_name TEXT NOT NULL,
                picks TEXT NOT NULL,
                PRIMARY KEY (week_id, employee_id, row_order)
            );
            CREATE INDEX IF NOT EXISTS idx_payroll_rows_employee
                ON payroll_rows(employee_id, week_id);
            CREATE INDEX IF NOT EXISTS idx_request_log_created_at
                ON request_log(created_at);
            """
        )
        staging_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(payroll_upload_staging)")
        }
        if "work_rows_json" not in staging_columns:
            connection.execute(
                "ALTER TABLE payroll_upload_staging "
                "ADD COLUMN work_rows_json TEXT NOT NULL DEFAULT '[]'"
            )
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute(
            "DELETE FROM payroll_upload_staging WHERE created_at < ?",
            (datetime.now(timezone.utc).timestamp() - 3600,),
        )
        connection.execute("DELETE FROM admins WHERE source = 'config'")
        now = datetime.now(timezone.utc).isoformat()
        connection.executemany(
            """
            INSERT INTO admins (telegram_id, source, added_by, added_at)
            VALUES (?, 'config', NULL, ?)
            ON CONFLICT(telegram_id) DO UPDATE SET source = 'config'
            """,
            [(telegram_id, now) for telegram_id in sorted(config_admin_ids or set())],
        )
    if DATABASE_PATH.exists():
        os.chmod(DATABASE_PATH, 0o600)


def log_request(telegram_id: int, action: str) -> None:
    with _connect() as connection:
        connection.execute(
            "INSERT INTO request_log (telegram_id, action, created_at) VALUES (?, ?, ?)",
            (telegram_id, action, datetime.now(timezone.utc).isoformat()),
        )


def save_payroll(
    week_id: str,
    headers: list[str],
    rows: list[tuple[str, str, list[str]]],
    expected_uploaded_at: str | None = None,
    check_version: bool = False,
    work_rows: list[tuple[str, str, str, str]] | None = None,
) -> None:
    with _connect() as connection:
        connection.execute("BEGIN IMMEDIATE")
        current = connection.execute(
            "SELECT uploaded_at FROM periods WHERE week_id = ?", (week_id,)
        ).fetchone()
        current_version = current["uploaded_at"] if current else None
        if check_version and current_version != expected_uploaded_at:
            raise ConcurrentPayrollUpdateError(
                "Эту неделю уже изменил другой администратор. Загрузка отменена; проверьте данные."
            )
        connection.execute(
            """
            INSERT INTO periods (week_id, headers_json, uploaded_at)
            VALUES (?, ?, ?)
            ON CONFLICT(week_id) DO UPDATE SET
                headers_json = excluded.headers_json,
                uploaded_at = excluded.uploaded_at
            """,
            (week_id, json.dumps(headers, ensure_ascii=False), datetime.now(timezone.utc).isoformat()),
        )
        connection.execute("DELETE FROM payroll_rows WHERE week_id = ?", (week_id,))
        connection.execute("DELETE FROM employee_work_rows WHERE week_id = ?", (week_id,))
        for normalized_name, display_name, values in rows:
            connection.execute(
                """
                INSERT INTO employees (normalized_name, display_name)
                VALUES (?, ?)
                ON CONFLICT(normalized_name) DO UPDATE SET display_name = excluded.display_name
                """,
                (normalized_name, display_name),
            )
            employee = connection.execute(
                "SELECT id FROM employees WHERE normalized_name = ?", (normalized_name,)
            ).fetchone()
            connection.execute(
                "INSERT INTO payroll_rows (week_id, employee_id, values_json) VALUES (?, ?, ?)",
                (week_id, employee["id"], json.dumps(values, ensure_ascii=False)),
            )
        employees_by_name = {
            row["normalized_name"]: row["id"]
            for row in connection.execute(
                """
                SELECT e.id, e.normalized_name
                FROM payroll_rows AS r
                JOIN employees AS e ON e.id = r.employee_id
                WHERE r.week_id = ?
                """,
                (week_id,),
            )
        }
        work_rows_by_employee: dict[str, int] = {}
        for normalized_name, display_name, work_name, picks in work_rows or []:
            employee_id = employees_by_name.get(normalized_name)
            if employee_id is None:
                raise ValueError(
                    f"Вкладка «Работа»: сотрудник «{display_name}» отсутствует "
                    "в расчетной таблице."
                )
            row_order = work_rows_by_employee.get(normalized_name, 0)
            connection.execute(
                """
                INSERT INTO employee_work_rows
                    (week_id, employee_id, row_order, work_name, picks)
                VALUES (?, ?, ?, ?, ?)
                """,
                (week_id, employee_id, row_order, work_name, picks),
            )
            work_rows_by_employee[normalized_name] = row_order + 1


def periods_for_employee(telegram_id: int) -> list[dict[str, Any]]:
    with _connect() as connection:
        employee = connection.execute(
            "SELECT id FROM employees WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if employee is None:
            return []
        periods = connection.execute(
            """
            SELECT p.week_id, p.headers_json, r.values_json
            FROM periods AS p
            LEFT JOIN payroll_rows AS r
                ON r.week_id = p.week_id AND r.employee_id = ?
            ORDER BY p.week_id DESC
            LIMIT 2
            """,
            (employee["id"],),
        ).fetchall()
    return [
        {
            "week_id": period["week_id"],
            "headers": json.loads(period["headers_json"]),
            "values": json.loads(period["values_json"]) if period["values_json"] else None,
        }
        for period in periods
    ]


def payroll_for_week(telegram_id: int, week_id: str) -> dict[str, Any] | None:
    with _connect() as connection:
        employee = connection.execute(
            "SELECT id FROM employees WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if employee is None:
            return None
        period = connection.execute(
            """
            SELECT p.week_id, p.headers_json, r.values_json
            FROM periods AS p
            LEFT JOIN payroll_rows AS r
                ON r.week_id = p.week_id AND r.employee_id = ?
            WHERE p.week_id = ?
            """,
            (employee["id"], week_id),
        ).fetchone()
    if period is None:
        return None
    return {
        "week_id": period["week_id"],
        "headers": json.loads(period["headers_json"]),
        "values": json.loads(period["values_json"]) if period["values_json"] else None,
    }


def period_exists(week_id: str) -> bool:
    with _connect() as connection:
        return (
            connection.execute(
                "SELECT 1 FROM periods WHERE week_id = ?", (week_id,)
            ).fetchone()
            is not None
        )


def period_uploaded_at(week_id: str) -> str | None:
    with _connect() as connection:
        period = connection.execute(
            "SELECT uploaded_at FROM periods WHERE week_id = ?", (week_id,)
        ).fetchone()
    return period["uploaded_at"] if period else None


def latest_period() -> str | None:
    with _connect() as connection:
        period = connection.execute(
            "SELECT week_id FROM periods ORDER BY week_id DESC LIMIT 1"
        ).fetchone()
    return period["week_id"] if period else None


def period_employee_display_names(week_id: str) -> dict[str, str]:
    with _connect() as connection:
        rows = connection.execute(
            """
            SELECT e.normalized_name, e.display_name
            FROM payroll_rows AS r
            JOIN employees AS e ON e.id = r.employee_id
            WHERE r.week_id = ?
            """,
            (week_id,),
        ).fetchall()
    return {row["normalized_name"]: row["display_name"] for row in rows}


def period_employee_values(week_id: str) -> dict[str, list[str]]:
    with _connect() as connection:
        rows = connection.execute(
            """
            SELECT e.normalized_name, r.values_json
            FROM payroll_rows AS r
            JOIN employees AS e ON e.id = r.employee_id
            WHERE r.week_id = ?
            """,
            (week_id,),
        ).fetchall()
    return {row["normalized_name"]: json.loads(row["values_json"]) for row in rows}


def period_headers(week_id: str) -> list[str] | None:
    with _connect() as connection:
        row = connection.execute(
            "SELECT headers_json FROM periods WHERE week_id = ?", (week_id,)
        ).fetchone()
    return json.loads(row["headers_json"]) if row else None


def work_rows_for_employee(telegram_id: int, week_id: str) -> list[dict[str, str]] | None:
    with _connect() as connection:
        employee = connection.execute(
            "SELECT id FROM employees WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if employee is None:
            return None
        rows = connection.execute(
            """
            SELECT work_name, picks
            FROM employee_work_rows
            WHERE week_id = ? AND employee_id = ?
            ORDER BY row_order
            """,
            (week_id, employee["id"]),
        ).fetchall()
    return [{"work_name": row["work_name"], "picks": row["picks"]} for row in rows]


def period_work_rows(week_id: str) -> dict[str, list[dict[str, str]]]:
    with _connect() as connection:
        rows = connection.execute(
            """
            SELECT e.normalized_name, w.work_name, w.picks
            FROM employee_work_rows AS w
            JOIN employees AS e ON e.id = w.employee_id
            WHERE w.week_id = ?
            ORDER BY e.normalized_name, w.row_order
            """,
            (week_id,),
        ).fetchall()
    result: dict[str, list[dict[str, str]]] = {}
    for row in rows:
        result.setdefault(row["normalized_name"], []).append(
            {"work_name": row["work_name"], "picks": row["picks"]}
        )
    return result


def stage_payroll_upload(
    token: str,
    admin_id: int,
    week_id: str,
    headers: list[str],
    rows: list[tuple[str, str, list[str]]],
    expected_uploaded_at: str | None,
    created_at: float,
    work_rows: list[tuple[str, str, str, str]] | None = None,
) -> None:
    with _connect() as connection:
        connection.execute("DELETE FROM payroll_upload_staging WHERE created_at < ?", (created_at - 3600,))
        connection.execute(
            """
            INSERT INTO payroll_upload_staging
                (token, admin_id, week_id, headers_json, rows_json, work_rows_json,
                 expected_uploaded_at, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                token,
                admin_id,
                week_id,
                json.dumps(headers, ensure_ascii=False),
                json.dumps(rows, ensure_ascii=False),
                json.dumps(work_rows or [], ensure_ascii=False),
                expected_uploaded_at,
                created_at,
            ),
        )


def get_staged_payroll_upload(token: str, admin_id: int, now: float) -> dict[str, Any] | None:
    with _connect() as connection:
        row = connection.execute(
            """
            SELECT week_id, headers_json, rows_json, work_rows_json, expected_uploaded_at
            FROM payroll_upload_staging
            WHERE token = ? AND admin_id = ? AND created_at >= ?
            """,
            (token, admin_id, now - 3600),
        ).fetchone()
    if row is None:
        return None
    return {
        "week_id": row["week_id"],
        "headers": json.loads(row["headers_json"]),
        "rows": json.loads(row["rows_json"]),
        "work_rows": json.loads(row["work_rows_json"]),
        "expected_uploaded_at": row["expected_uploaded_at"],
    }


def delete_staged_payroll_upload(token: str, admin_id: int) -> None:
    with _connect() as connection:
        connection.execute(
            "DELETE FROM payroll_upload_staging WHERE token = ? AND admin_id = ?",
            (token, admin_id),
        )


def has_staged_payroll_upload(token: str, admin_id: int, now: float) -> bool:
    with _connect() as connection:
        return (
            connection.execute(
                """
                SELECT 1 FROM payroll_upload_staging
                WHERE token = ? AND admin_id = ? AND created_at >= ?
                """,
                (token, admin_id, now - 3600),
            ).fetchone()
            is not None
        )


def employee_by_name(normalized_name: str) -> sqlite3.Row | None:
    with _connect() as connection:
        return connection.execute(
            "SELECT id, password_hash, telegram_id FROM employees WHERE normalized_name = ?",
            (normalized_name,),
        ).fetchone()


def is_telegram_linked(telegram_id: int) -> bool:
    with _connect() as connection:
        return (
            connection.execute(
                "SELECT 1 FROM employees WHERE telegram_id = ?", (telegram_id,)
            ).fetchone()
            is not None
        )


def employee_names_without_passwords() -> list[sqlite3.Row]:
    with _connect() as connection:
        return connection.execute(
            """
            SELECT id, normalized_name, display_name FROM employees
            WHERE password_hash IS NULL ORDER BY display_name
            """
        ).fetchall()


def employee_directory() -> list[sqlite3.Row]:
    with _connect() as connection:
        return connection.execute(
            """
            SELECT display_name, password_hash IS NOT NULL AS has_password,
                   telegram_id IS NOT NULL AS is_linked
            FROM employees ORDER BY display_name
            """
        ).fetchall()


def employee_count() -> int:
    with _connect() as connection:
        return connection.execute("SELECT COUNT(*) FROM employees").fetchone()[0]


def linked_employee_count() -> int:
    with _connect() as connection:
        return connection.execute(
            "SELECT COUNT(*) FROM employees WHERE telegram_id IS NOT NULL"
        ).fetchone()[0]


def set_password(employee_id: int, password_hash: bytes) -> None:
    with _connect() as connection:
        connection.execute(
            "UPDATE employees SET password_hash = ? WHERE id = ?",
            (password_hash, employee_id),
        )


def set_passwords(password_hashes: list[tuple[int, bytes]]) -> None:
    with _connect() as connection:
        cursor = connection.executemany(
            "UPDATE employees SET password_hash = ? WHERE id = ? AND password_hash IS NULL",
            [(password_hash, employee_id) for employee_id, password_hash in password_hashes],
        )
        if cursor.rowcount != len(password_hashes):
            raise RuntimeError("Набор новых паролей изменился во время операции.")


def clear_passwords_if_match(password_hashes: list[tuple[int, bytes]]) -> None:
    with _connect() as connection:
        connection.executemany(
            "UPDATE employees SET password_hash = NULL WHERE id = ? AND password_hash = ?",
            [(employee_id, password_hash) for employee_id, password_hash in password_hashes],
        )


def restore_password(employee_id: int, password_hash: bytes | None) -> None:
    with _connect() as connection:
        connection.execute(
            "UPDATE employees SET password_hash = ? WHERE id = ?",
            (password_hash, employee_id),
        )


def unlink_telegram_employee(employee_id: int) -> None:
    with _connect() as connection:
        connection.execute("BEGIN IMMEDIATE")
        connection.execute(
            "UPDATE employees SET telegram_id = NULL WHERE id = ?", (employee_id,)
        )


def link_telegram_account(employee_id: int, telegram_id: int) -> bool:
    with _connect() as connection:
        connection.execute("BEGIN IMMEDIATE")
        existing = connection.execute(
            "SELECT telegram_id FROM employees WHERE id = ?", (employee_id,)
        ).fetchone()
        linked_to_another_employee = connection.execute(
            "SELECT id FROM employees WHERE telegram_id = ? AND id != ?",
            (telegram_id, employee_id),
        ).fetchone()
        if existing is None or (
            existing["telegram_id"] is not None and existing["telegram_id"] != telegram_id
        ) or linked_to_another_employee is not None:
            return False
        connection.execute(
            "UPDATE employees SET telegram_id = ? WHERE id = ?", (telegram_id, employee_id)
        )
        return True


def unlink_telegram_account(normalized_name: str) -> bool:
    with _connect() as connection:
        result = connection.execute(
            "UPDATE employees SET telegram_id = NULL WHERE normalized_name = ?",
            (normalized_name,),
        )
        return result.rowcount > 0


def unlink_telegram_id(telegram_id: int) -> None:
    with _connect() as connection:
        connection.execute(
            "UPDATE employees SET telegram_id = NULL WHERE telegram_id = ?", (telegram_id,)
        )


def login_lockout(telegram_id: int, now: float) -> float:
    with _connect() as connection:
        attempt = connection.execute(
            "SELECT locked_until FROM login_attempts WHERE telegram_id = ?",
            (telegram_id,),
        ).fetchone()
    if attempt is None:
        return 0
    if attempt["locked_until"] > now:
        return attempt["locked_until"]
    return 0


def record_login_failure(
    telegram_id: int, now: float, max_failures: int, window_seconds: int, lock_seconds: int
) -> float:
    with _connect() as connection:
        connection.execute("BEGIN IMMEDIATE")
        attempt = connection.execute(
            "SELECT failures, window_started, locked_until FROM login_attempts WHERE telegram_id = ?",
            (telegram_id,),
        ).fetchone()
        if attempt is None or now - attempt["window_started"] >= window_seconds:
            failures = 1
            window_started = now
        else:
            failures = attempt["failures"] + 1
            window_started = attempt["window_started"]
        locked_until = now + lock_seconds if failures >= max_failures else 0
        connection.execute(
            """
            INSERT INTO login_attempts (telegram_id, failures, window_started, locked_until)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(telegram_id) DO UPDATE SET
                failures = excluded.failures,
                window_started = excluded.window_started,
                locked_until = excluded.locked_until
            """,
            (telegram_id, failures, window_started, locked_until),
        )
    return locked_until


def clear_login_failures(telegram_id: int) -> None:
    with _connect() as connection:
        connection.execute("DELETE FROM login_attempts WHERE telegram_id = ?", (telegram_id,))


def is_admin(telegram_id: int) -> bool:
    with _connect() as connection:
        return (
            connection.execute(
                "SELECT 1 FROM admins WHERE telegram_id = ?", (telegram_id,)
            ).fetchone()
            is not None
        )


def list_admins() -> list[sqlite3.Row]:
    with _connect() as connection:
        return connection.execute(
            "SELECT telegram_id, source FROM admins ORDER BY telegram_id"
        ).fetchall()


def add_admin(telegram_id: int, added_by: int) -> bool:
    if telegram_id <= 0:
        return False
    with _connect() as connection:
        result = connection.execute(
            """
            INSERT OR IGNORE INTO admins (telegram_id, source, added_by, added_at)
            VALUES (?, 'database', ?, ?)
            """,
            (telegram_id, added_by, datetime.now(timezone.utc).isoformat()),
        )
        return result.rowcount == 1


def remove_admin(telegram_id: int, actor_id: int) -> str:
    with _connect() as connection:
        connection.execute("BEGIN IMMEDIATE")
        admin = connection.execute(
            "SELECT source FROM admins WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if admin is None:
            return "not_found"
        if admin["source"] == "config":
            return "protected"
        if telegram_id == actor_id:
            return "self"
        count = connection.execute("SELECT COUNT(*) FROM admins").fetchone()[0]
        if count <= 1:
            return "last_admin"
        connection.execute("DELETE FROM admins WHERE telegram_id = ?", (telegram_id,))
        return "removed"


def load_fsm(storage_key: str) -> tuple[str | None, dict[str, Any]]:
    with _connect() as connection:
        row = connection.execute(
            "SELECT state, data_json FROM fsm_storage WHERE storage_key = ?",
            (storage_key,),
        ).fetchone()
    if row is None:
        return None, {}
    data = json.loads(row["data_json"])
    if not isinstance(data, dict):
        raise ValueError("Invalid persisted FSM data")
    return row["state"], data


def save_fsm_state(storage_key: str, state: str | None) -> None:
    with _connect() as connection:
        connection.execute(
            """
            INSERT INTO fsm_storage (storage_key, state, data_json)
            VALUES (?, ?, '{}')
            ON CONFLICT(storage_key) DO UPDATE SET state = excluded.state
            """,
            (storage_key, state),
        )


def save_fsm_data(storage_key: str, data: dict[str, Any]) -> None:
    with _connect() as connection:
        connection.execute(
            """
            INSERT INTO fsm_storage (storage_key, state, data_json)
            VALUES (?, NULL, ?)
            ON CONFLICT(storage_key) DO UPDATE SET data_json = excluded.data_json
            """,
            (storage_key, json.dumps(data, ensure_ascii=False)),
        )
