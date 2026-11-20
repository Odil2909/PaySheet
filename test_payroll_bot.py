from __future__ import annotations

import asyncio
import csv
import io
import json
import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock

os.environ.setdefault("BOT_TOKEN", "test-token")
os.environ.setdefault("OWNER_IDS", "1")

import bcrypt
from aiogram.fsm.storage.base import StorageKey
from openpyxl import Workbook

import database
from payroll_bot import (
    Login,
    cancel,
    format_amount,
    is_admin,
    is_private_admin_callback,
    keyboard,
    parse_week_id,
    parse_workbook,
    parse_workbook_with_work,
    payroll_diff_file,
    render_payroll,
    render_work,
    roster_review,
    start,
    upload_start,
    week_id_from_headers,
    week_label,
    week_picker,
    work_review_changes,
)
from storage import SQLiteFSMStorage


class PayrollBotTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_database_path = database.DATABASE_PATH
        database.DATABASE_PATH = Path(self.temp_dir.name) / "test.sqlite3"
        database.initialize({100})

    def tearDown(self) -> None:
        database.DATABASE_PATH = self.original_database_path
        self.temp_dir.cleanup()

    def test_owner_keyboard_has_management_buttons(self) -> None:
        employee_buttons = {button.text for row in keyboard().keyboard for button in row}
        owner_buttons = {button.text for row in keyboard(owner=True).keyboard for button in row}
        self.assertIn("🔐 Войти", employee_buttons)
        self.assertIn("📅 Текущая неделя", employee_buttons)
        self.assertIn("📦 Работа и пики", employee_buttons)
        self.assertNotIn("📤 Загрузить Excel", employee_buttons)
        self.assertIn("📤 Загрузить Excel", owner_buttons)
        self.assertIn("🔑 Пароли сотрудников", owner_buttons)
        self.assertIn("🔄 Сброс пароля", owner_buttons)
        self.assertIn("⚙️ Администрирование", owner_buttons)

    def test_money_uses_decimal_rounding_and_grouping(self) -> None:
        self.assertEqual(format_amount("6170.866200000001"), "6 170.87")
        self.assertEqual(format_amount("7896.155700000002"), "7 896.16")
        self.assertEqual(format_amount(0), "0.00")
        self.assertEqual(format_amount(None), "—")
        self.assertEqual(format_amount("1.234,50"), "1 234.50")

    def test_week_labels_and_excel_week_detection(self) -> None:
        headers = [
            "ФИО",
            "2026-11-09",
            "2026-11-10",
            "2026-11-11",
            "2026-11-12",
            "2026-11-13",
            "2026-11-14",
            "2026-11-15",
            "Штрафы",
            "Итого",
        ]
        self.assertEqual(week_id_from_headers(headers), "2026-W46")
        self.assertEqual(week_label("2026-W46"), "9–15 ноября 2026")
        headers[2] = "2026-11-12"
        with self.assertRaisesRegex(ValueError, "подряд"):
            week_id_from_headers(headers)
        self.assertEqual(week_label("2027-W05"), "1–7 февраля 2027")
        self.assertEqual(week_label("2026-W53"), "28 декабря 2026 – 3 января 2027")

    def test_week_picker_has_recent_weeks_and_custom_option(self) -> None:
        picker = week_picker()
        week_ids = [
            row[0].callback_data.removeprefix("upload_week:")
            for row in picker.inline_keyboard
            if row[0].callback_data.startswith("upload_week:20")
        ]
        self.assertEqual(len(week_ids), 4)
        for week_id in week_ids:
            self.assertEqual(parse_week_id(week_id), week_id)
        self.assertIn(
            "upload_week:custom",
            [row[0].callback_data for row in picker.inline_keyboard],
        )

    @staticmethod
    def make_xlsx(rows: list[list[object]]) -> bytes:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "ФИО",
                "Пн",
                "Вт",
                "Ср",
                "Чт",
                "Пт",
                "Сб",
                "Вс",
                "Штрафы",
                "Доплаты/удержания",
                "Итого",
            ]
        )
        for row in rows:
            sheet.append(row)
        output = io.BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_parse_payroll_xlsx(self) -> None:
        headers, rows = parse_workbook(
            self.make_xlsx([["Иванов Иван", 1, 2, 3, 4, 5, 6, 7, 100, -20, 800]])
        )
        self.assertEqual(headers[0], "ФИО")
        self.assertEqual(rows[0][0], "иванов иван")
        self.assertEqual(rows[0][1], "Иванов Иван")
        self.assertEqual(rows[0][2], ["1", "2", "3", "4", "5", "6", "7", "100", "-20", "800"])

    def test_parse_work_sheet_with_multiple_service_rows(self) -> None:
        workbook = Workbook()
        payroll = workbook.active
        payroll.append(
            [
                "ФИО",
                "Пн",
                "Вт",
                "Ср",
                "Чт",
                "Пт",
                "Сб",
                "Вс",
                "Штрафы",
                "Итого",
            ]
        )
        payroll.append(["Иванов Иван", 1, 2, 3, 4, 5, 6, 7, 0, 28])
        work = workbook.create_sheet("Работа")
        work.append(["ФИО", "Работа", "Пики, шт."])
        work.append(["Иванов Иван", "Сборка", 120])
        work.append(["Иванов Иван", "Упаковка", 0])
        output = io.BytesIO()
        workbook.save(output)

        headers, rows, work_rows = parse_workbook_with_work(output.getvalue())

        self.assertEqual(headers[0], "ФИО")
        self.assertEqual(len(rows), 1)
        self.assertEqual(
            work_rows,
            [
                ("иванов иван", "Иванов Иван", "Сборка", "120"),
                ("иванов иван", "Иванов Иван", "Упаковка", "0"),
            ],
        )

    def test_work_sheet_rejects_employee_missing_from_payroll(self) -> None:
        workbook = Workbook()
        payroll = workbook.active
        payroll.append(
            [
                "ФИО",
                "Пн",
                "Вт",
                "Ср",
                "Чт",
                "Пт",
                "Сб",
                "Вс",
                "Штрафы",
                "Итого",
            ]
        )
        payroll.append(["Иванов Иван", 1, 2, 3, 4, 5, 6, 7, 0, 28])
        work = workbook.create_sheet("Работа")
        work.append(["ФИО", "Работа", "Пики, шт."])
        work.append(["Петров Петр", "Сборка", 120])
        output = io.BytesIO()
        workbook.save(output)

        with self.assertRaisesRegex(ValueError, "отсутствует в расчетной таблице"):
            parse_workbook_with_work(output.getvalue())

    def test_parser_finds_payroll_table_on_a_non_active_sheet(self) -> None:
        workbook = Workbook()
        workbook.active.title = "Сводка"
        workbook.active["A1"] = "служебная таблица"
        sheet = workbook.create_sheet("Расчёт")
        sheet.append(
            [
                "ФИО",
                date(2026, 11, 9),
                date(2026, 11, 10),
                date(2026, 11, 11),
                date(2026, 11, 12),
                date(2026, 11, 13),
                date(2026, 11, 14),
                date(2026, 11, 15),
                "Штрафы",
                "Итого",
            ]
        )
        sheet.append(["Иванов Иван", 0, 0, 0, 0, 0, 0, 0, 0, 0])
        output = io.BytesIO()
        workbook.save(output)
        headers, rows = parse_workbook(output.getvalue())
        self.assertEqual(headers[0], "ФИО")
        self.assertEqual(rows[0][1], "Иванов Иван")
        rendered = render_payroll(
            {"week_id": "2026-W46", "headers": headers, "values": rows[0][2]},
            "текущую неделю",
        )
        self.assertEqual(rendered.count("0.00"), 9)

    def test_parse_sample_format_without_additions_column(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "ФИО",
                date(2026, 11, 9),
                date(2026, 11, 10),
                date(2026, 11, 11),
                date(2026, 11, 12),
                date(2026, 11, 13),
                date(2026, 11, 14),
                date(2026, 11, 15),
                "Штрафы",
                "Итого",
            ]
        )
        sheet.append(["Иванов Иван", 1, 2, 3, 4, 5, 6, 7, 100, 800])
        output = io.BytesIO()
        workbook.save(output)

        headers, rows = parse_workbook(output.getvalue())
        self.assertEqual(len(headers), 10)
        self.assertEqual(headers[1], "2026-11-09")
        rendered = render_payroll(
            {"week_id": "2026-W46", "headers": headers, "values": rows[0][2]},
            "текущую неделю",
        )
        self.assertIn("Неделя: 9–15 ноября 2026", rendered)
        self.assertIn("📆 09.11 — 1.00", rendered)
        self.assertNotIn("Доплаты", rendered)
        self.assertIn("ИТОГО: 800.00", rendered)

    def test_combined_and_separate_adjustment_columns_are_formatted(self) -> None:
        record = {
            "week_id": "2026-W46",
            "headers": [
                "ФИО",
                *[f"День {index}" for index in range(1, 8)],
                "Штрафы",
                "Доплаты",
                "Удержания",
                "Итого",
            ],
            "values": [
                "0",
                "0",
                "0",
                "0",
                "0",
                "0",
                "0",
                "6170.866200000001",
                "7896.155700000002",
                "0",
                "14067.021900000003",
            ],
        }
        rendered = render_payroll(record, "текущую неделю")
        self.assertIn("⚠️ Штрафы: 6 170.87", rendered)
        self.assertIn("➕ Доплаты: 7 896.16", rendered)
        self.assertIn("➖ Удержания: 0.00", rendered)
        self.assertIn("ИТОГО: 14 067.02", rendered)
        self.assertEqual(record["values"][0], "0")

    def test_duplicate_employee_names_are_rejected(self) -> None:
        content = self.make_xlsx(
            [
                ["Иванов Иван", 1, 2, 3, 4, 5, 6, 7, 0, 0, 800],
                ["  Иванов   Иван ", 1, 2, 3, 4, 5, 6, 7, 0, 0, 800],
            ]
        )
        with self.assertRaisesRegex(ValueError, "повторяется"):
            parse_workbook(content)

    def test_weekly_updates_preserve_passwords_and_history(self) -> None:
        password_hash = bcrypt.hashpw(b"secret", bcrypt.gensalt())
        first_rows = [
            ("иванов иван", "Иванов Иван", ["1"] * 10),
            ("петрова анна", "Петрова Анна", ["2"] * 10),
        ]
        database.save_payroll("2026-W46", ["ФИО"] + ["День"] * 10, first_rows)
        employee = database.employee_by_name("иванов иван")
        database.set_password(employee["id"], password_hash)
        self.assertTrue(database.link_telegram_account(employee["id"], 456))

        database.save_payroll(
            "2026-W47",
            ["ФИО"] + ["День"] * 10,
            [("иванов иван", "Иванов Иван", ["3"] * 10)],
        )
        updated = database.employee_by_name("иванов иван")
        self.assertEqual(bytes(updated["password_hash"]), password_hash)
        self.assertTrue(database.is_telegram_linked(456))

        periods = database.periods_for_employee(456)
        self.assertEqual([period["week_id"] for period in periods], ["2026-W47", "2026-W46"])
        self.assertEqual(periods[0]["values"], ["3"] * 10)

        other_employee = database.employee_by_name("петрова анна")
        self.assertTrue(database.link_telegram_account(other_employee["id"], 789))
        other_periods = database.periods_for_employee(789)
        self.assertEqual([period["week_id"] for period in other_periods], ["2026-W47", "2026-W46"])
        self.assertIsNone(other_periods[0]["values"])
        self.assertEqual(other_periods[1]["values"], ["2"] * 10)

    def test_existing_week_replacement_is_guarded_against_concurrent_admin(self) -> None:
        headers = ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"]
        rows = [("иванов иван", "Иванов Иван", ["1"] * 9)]
        database.save_payroll("2026-W46", headers, rows)
        version = database.period_uploaded_at("2026-W46")
        database.save_payroll(
            "2026-W46", headers, [("иванов иван", "Иванов Иван", ["2"] * 9)]
        )
        with self.assertRaises(database.ConcurrentPayrollUpdateError):
            database.save_payroll(
                "2026-W46",
                headers,
                [("иванов иван", "Иванов Иван", ["3"] * 9)],
                expected_uploaded_at=version,
                check_version=True,
            )
        self.assertEqual(
            database.period_employee_values("2026-W46")["иванов иван"], ["2"] * 9
        )

    def test_staged_upload_is_private_to_admin_and_expires(self) -> None:
        work_rows = [("имя", "Имя", "Сборка", "15")]
        database.stage_payroll_upload(
            "a" * 32,
            100,
            "2026-W46",
            ["ФИО"],
            [("имя", "Имя", ["0"])],
            None,
            1000,
            work_rows=work_rows,
        )
        self.assertIsNone(database.get_staged_payroll_upload("a" * 32, 200, 1001))
        self.assertIsNone(database.get_staged_payroll_upload("a" * 32, 100, 4601))
        staged = database.get_staged_payroll_upload("a" * 32, 100, 1001)
        self.assertIsNotNone(staged)
        self.assertEqual(staged["work_rows"], [list(row) for row in work_rows])

    def test_work_rows_are_isolated_by_linked_employee_and_week(self) -> None:
        rows = [
            ("иванов иван", "Иванов Иван", ["0"] * 9),
            ("петров петр", "Петров Петр", ["0"] * 9),
        ]
        work_rows = [
            ("иванов иван", "Иванов Иван", "Сборка", "12"),
            ("иванов иван", "Иванов Иван", "Упаковка", "0"),
        ]
        database.save_payroll(
            "2026-W46",
            ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"],
            rows,
            work_rows=work_rows,
        )
        employee = database.employee_by_name("иванов иван")
        self.assertTrue(database.link_telegram_account(employee["id"], 456))
        self.assertEqual(
            database.work_rows_for_employee(456, "2026-W46"),
            [
                {"work_name": "Сборка", "picks": "12"},
                {"work_name": "Упаковка", "picks": "0"},
            ],
        )
        self.assertEqual(
            database.work_rows_for_employee(456, "2026-W47"),
            [],
        )
        other_employee = database.employee_by_name("петров петр")
        self.assertTrue(database.link_telegram_account(other_employee["id"], 789))
        self.assertEqual(database.work_rows_for_employee(789, "2026-W46"), [])
        self.assertEqual(
            render_work("2026-W46", database.work_rows_for_employee(456, "2026-W46")),
            (
                "📦 Работа и пики\n📅 Неделя: 9–15 ноября 2026\n\n"
                "━━━━━━━━━━━━━━\n\n"
                "• Сборка — 12 шт.\n• Упаковка — 0 шт.\n\n"
                "━━━━━━━━━━━━━━\n\nИтого пиков: 12 шт."
            ),
        )

    def test_roster_review_lists_exact_changes_and_detects_spelling_swap(self) -> None:
        headers = ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"]
        database.save_payroll(
            "2026-W46",
            headers,
            [
                ("иванов иван", "Иванов Иван", ["1"] * 9),
                ("петров петр", "Петров Петр", ["2"] * 9),
            ],
        )
        removed, added, changed, possible_renames = roster_review(
            "2026-W47",
            [
                ("иванов иван", "Иванов Иван", ["3"] * 9),
                ("сидоров сергей", "Сидоров Сергей", ["4"] * 9),
            ],
        )
        self.assertEqual(removed, ["Петров Петр"])
        self.assertEqual(added, ["Сидоров Сергей"])
        self.assertEqual(changed, [])
        self.assertEqual(possible_renames, [])

        database.save_payroll(
            "2026-W47",
            headers,
            [
                ("иванов иван", "Иванов Иван", ["3"] * 9),
                ("петров петр", "Петров Петр", ["2"] * 9),
            ],
        )
        _, _, _, possible_renames = roster_review(
            "2026-W47",
            [
                ("иванов иван", "Иванов Иван", ["3"] * 9),
                ("петрова петр", "Петрова Петр", ["2"] * 9),
            ],
        )
        self.assertEqual(possible_renames, [("Петров Петр", "Петрова Петр")])

    def test_payroll_diff_export_shows_exact_old_and_new_rows(self) -> None:
        headers = ["ФИО"] + [f"День {index}" for index in range(1, 8)] + [
            "Штрафы",
            "Итого",
        ]
        previous_values = ["1", "2", "3", "4", "5", "6", "7", "10", "38"]
        current_values = ["1", "2", "3", "4", "5", "6", "7", "12", "40"]
        database.save_payroll(
            "2026-W46",
            headers,
            [("иванов иван", "Иванов Иван", previous_values)],
        )
        review_file = payroll_diff_file(
            "2026-W46",
            headers,
            [("иванов иван", "Иванов Иван", current_values)],
            [],
            [],
            ["Иванов Иван"],
        )
        exported = list(csv.DictReader(io.StringIO(review_file.data.decode("utf-8-sig"))))
        self.assertEqual(exported[0]["Изменение"], "Изменились данные строки")
        self.assertEqual(json.loads(exported[0]["До (поля Excel)"])["Штрафы"], "10")
        self.assertEqual(json.loads(exported[0]["После (поля Excel)"])["Штрафы"], "12")

    def test_work_only_changes_are_included_in_upload_review(self) -> None:
        headers = ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"]
        rows = [("иванов иван", "Иванов Иван", ["1"] * 9)]
        database.save_payroll(
            "2026-W46",
            headers,
            rows,
            work_rows=[("иванов иван", "Иванов Иван", "Сборка", "10")],
        )
        incoming_work = [("иванов иван", "Иванов Иван", "Сборка", "15")]
        changed_work = work_review_changes("2026-W46", rows, incoming_work)
        review_file = payroll_diff_file(
            "2026-W46", headers, rows, [], [], [], incoming_work, changed_work
        )
        exported = list(csv.DictReader(io.StringIO(review_file.data.decode("utf-8-sig"))))

        self.assertEqual(changed_work, ["Иванов Иван"])
        self.assertEqual(exported[0]["Изменение"], "Изменилась работа и количество пиков")
        self.assertEqual(
            json.loads(exported[0]["До (работа и пики)"])[0]["picks"],
            "10",
        )
        self.assertEqual(
            json.loads(exported[0]["После (работа и пики)"])[0]["picks"],
            "15",
        )

    def test_fsm_state_and_data_survive_storage_recreation(self) -> None:
        key = StorageKey(bot_id=1, chat_id=200, user_id=200)

        async def persist_and_reload() -> tuple[str | None, dict[str, object]]:
            first = SQLiteFSMStorage()
            await first.set_state(key, Login.name)
            await first.set_data(key, {"employee_name": "Иванов Иван"})
            second = SQLiteFSMStorage()
            state = await second.get_state(key)
            data = await second.get_data(key)
            await second.set_state(key, None)
            await second.set_data(key, {})
            return state, data

        state, data = asyncio.run(persist_and_reload())
        self.assertEqual(state, Login.name.state)
        self.assertEqual(data, {"employee_name": "Иванов Иван"})

    def test_admin_check_requires_private_chat_and_database_admin(self) -> None:
        private_admin = SimpleNamespace(
            from_user=SimpleNamespace(id=100),
            chat=SimpleNamespace(type="private"),
        )
        group_admin = SimpleNamespace(
            from_user=SimpleNamespace(id=100),
            chat=SimpleNamespace(type="group"),
        )
        non_admin = SimpleNamespace(
            from_user=SimpleNamespace(id=200),
            chat=SimpleNamespace(type="private"),
        )
        self.assertTrue(is_admin(private_admin))
        self.assertFalse(is_admin(group_admin))
        self.assertFalse(is_admin(non_admin))
        self.assertTrue(
            is_private_admin_callback(
                SimpleNamespace(
                    from_user=SimpleNamespace(id=100),
                    message=SimpleNamespace(chat=SimpleNamespace(type="private")),
                )
            )
        )
        self.assertFalse(
            is_private_admin_callback(
                SimpleNamespace(
                    from_user=SimpleNamespace(id=100),
                    message=SimpleNamespace(chat=SimpleNamespace(type="group")),
                )
            )
        )

    def test_start_and_cancel_restore_the_role_menu(self) -> None:
        async def run() -> None:
            admin_message = SimpleNamespace(
                from_user=SimpleNamespace(id=100),
                chat=SimpleNamespace(type="private"),
                answer=AsyncMock(),
            )
            state = AsyncMock()
            state.get_data.return_value = {}
            await start(admin_message, state)
            admin_start_markup = admin_message.answer.await_args.kwargs["reply_markup"]
            await cancel(admin_message, state)
            admin_cancel_markup = admin_message.answer.await_args.kwargs["reply_markup"]
            self.assertIn(
                "📤 Загрузить Excel",
                {button.text for row in admin_start_markup.keyboard for button in row},
            )
            self.assertIn(
                "👥 Сотрудники",
                {button.text for row in admin_cancel_markup.keyboard for button in row},
            )

            employee_message = SimpleNamespace(
                from_user=SimpleNamespace(id=200),
                chat=SimpleNamespace(type="private"),
                answer=AsyncMock(),
            )
            await start(employee_message, state)
            employee_markup = employee_message.answer.await_args.kwargs["reply_markup"]
            self.assertNotIn(
                "📤 Загрузить Excel",
                {button.text for row in employee_markup.keyboard for button in row},
            )
        asyncio.run(run())

    def test_non_admin_upload_command_is_rejected(self) -> None:
        async def run() -> None:
            message = SimpleNamespace(
                from_user=SimpleNamespace(id=200),
                chat=SimpleNamespace(type="private"),
                answer=AsyncMock(),
            )
            state = AsyncMock()
            command = SimpleNamespace(args="2026-W46")
            await upload_start(message, command, state)
            message.answer.assert_awaited_once()
            self.assertIn("только владельцу", message.answer.await_args.args[0])
            state.clear.assert_not_awaited()

        asyncio.run(run())

    def test_exact_week_query_and_employee_isolation(self) -> None:
        database.save_payroll(
            "2026-W46",
            ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"],
            [("сотрудник а", "Сотрудник А", ["1"] * 9)],
        )
        database.save_payroll(
            "2026-W47",
            ["ФИО"] + ["День"] * 7 + ["Штрафы", "Итого"],
            [("сотрудник б", "Сотрудник Б", ["2"] * 9)],
        )
        first = database.employee_by_name("сотрудник а")
        second = database.employee_by_name("сотрудник б")
        self.assertTrue(database.link_telegram_account(first["id"], 456))
        self.assertTrue(database.link_telegram_account(second["id"], 789))
        self.assertEqual(database.payroll_for_week(456, "2026-W46")["values"], ["1"] * 9)
        self.assertIsNone(database.payroll_for_week(456, "2026-W47")["values"])
        self.assertIsNone(database.payroll_for_week(456, "2026-W45"))

    def test_admin_permissions_persist_and_protect_configured_admins(self) -> None:
        self.assertTrue(database.is_admin(100))
        self.assertTrue(database.add_admin(200, 100))
        self.assertFalse(database.add_admin(200, 100))
        self.assertTrue(database.is_admin(200))
        self.assertEqual(database.remove_admin(100, 200), "protected")
        self.assertEqual(database.remove_admin(200, 200), "self")
        database.initialize({100})
        self.assertTrue(database.is_admin(200))
        self.assertEqual(database.remove_admin(200, 100), "removed")
        self.assertFalse(database.is_admin(200))

    def test_database_permissions_are_private(self) -> None:
        mode = self.temp_dir.name
        database.initialize({100})
        path = Path(mode) / "test.sqlite3"
        self.assertEqual(path.stat().st_mode & 0o777, 0o600)

    def test_login_lockout_after_five_failures(self) -> None:
        now = 1_800_000_000.0
        self.assertEqual(database.login_lockout(789, now), 0)
        for _ in range(4):
            self.assertEqual(database.record_login_failure(789, now, 5, 900, 900), 0)
        locked_until = database.record_login_failure(789, now, 5, 900, 900)
        self.assertEqual(locked_until, now + 900)
        self.assertEqual(database.login_lockout(789, now + 10), locked_until)


if __name__ == "__main__":
    unittest.main()
