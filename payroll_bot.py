from __future__ import annotations

import asyncio
import csv
import difflib
import io
import json
import logging
import os
import re
import secrets
import sqlite3
import time
import zipfile
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any
from xml.etree import ElementTree

import bcrypt
from aiogram import Bot, Dispatcher, F, Router
from aiogram.exceptions import TelegramAPIError
from aiogram.filters import Command, CommandObject
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    ErrorEvent,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

import database
from storage import SQLiteFSMStorage


load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()


def configured_admin_ids() -> set[int]:
    try:
        admin_ids = {
            int(value.strip())
            for variable in ("OWNER_IDS", "ADMIN_IDS")
            for value in os.getenv(variable, "").split(",")
            if value.strip()
        }
    except ValueError as exc:
        raise RuntimeError("OWNER_IDS и ADMIN_IDS должны содержать числовые Telegram ID") from exc
    if any(telegram_id <= 0 for telegram_id in admin_ids):
        raise RuntimeError("Telegram ID администратора должен быть положительным числом")
    return admin_ids


CONFIG_ADMIN_IDS = configured_admin_ids()
if not BOT_TOKEN:
    raise RuntimeError("Укажите токен бота в BOT_TOKEN в файле .env")
if not CONFIG_ADMIN_IDS:
    raise RuntimeError("Укажите хотя бы одного администратора в OWNER_IDS или ADMIN_IDS в .env")

MAX_LOGIN_FAILURES = 5
LOGIN_WINDOW_SECONDS = 15 * 60
LOGIN_LOCK_SECONDS = 15 * 60
MAX_XLSX_BYTES = 10 * 1024 * 1024
MAX_XLSX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024
WEEK_ID_PATTERN = re.compile(r"^\d{4}-W(?:0[1-9]|[1-4]\d|5[0-3])$")
logger = logging.getLogger(__name__)
MENU_BUTTONS = {
    "📅 Текущая неделя",
    "📅 Предыдущая неделя",
    "📦 Работа и пики",
    "🔐 Войти",
    "🚪 Выйти",
    "📤 Загрузить Excel",
    "👥 Сотрудники",
    "🔑 Пароли сотрудников",
    "🔄 Сброс пароля",
    "⚙️ Администрирование",
    "Текущая неделя",
    "Предыдущая неделя",
    "Войти",
    "Выйти",
    "Загрузить Excel",
    "Пароли сотрудникам",
    "Сбросить пароль",
    "Сотрудники",
    "Администрирование",
}
RU_MONTHS = (
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
)
STARTED_AT = time.monotonic()

router = Router()
dispatcher = Dispatcher(storage=SQLiteFSMStorage())
dispatcher.include_router(router)


class Login(StatesGroup):
    name = State()
    password = State()


class Upload(StatesGroup):
    week = State()
    confirm_replace = State()
    document = State()
    review = State()


class ResetPassword(StatesGroup):
    employee_name = State()


class AdminManagement(StatesGroup):
    add_admin_id = State()


def normalize_name(value: str) -> str:
    return " ".join(value.split()).casefold()


def keyboard(owner: bool = False) -> ReplyKeyboardMarkup:
    rows = [
        [
            KeyboardButton(text="📅 Текущая неделя"),
            KeyboardButton(text="📅 Предыдущая неделя"),
        ],
        [KeyboardButton(text="📦 Работа и пики")],
        [KeyboardButton(text="🔐 Войти"), KeyboardButton(text="🚪 Выйти")],
    ]
    if owner:
        rows.extend(
            [
                [KeyboardButton(text="📤 Загрузить Excel"), KeyboardButton(text="👥 Сотрудники")],
                [
                    KeyboardButton(text="🔑 Пароли сотрудников"),
                    KeyboardButton(text="🔄 Сброс пароля"),
                ],
                [KeyboardButton(text="⚙️ Администрирование")],
            ]
        )
    return ReplyKeyboardMarkup(
        keyboard=rows,
        resize_keyboard=True,
    )


def week_picker() -> InlineKeyboardMarkup:
    today = datetime.now().astimezone().date()
    current_monday = today - timedelta(days=today.isoweekday() - 1)
    buttons = []
    for offset in range(4):
        week_start = current_monday - timedelta(weeks=offset)
        iso_year, iso_week, _ = week_start.isocalendar()
        week_id = f"{iso_year}-W{iso_week:02d}"
        label = f"Текущая — {week_id}" if offset == 0 else week_id
        buttons.append([InlineKeyboardButton(text=label, callback_data=f"upload_week:{week_id}")])
    buttons.extend(
        [
            [InlineKeyboardButton(text="Указать другую неделю", callback_data="upload_week:custom")],
            [InlineKeyboardButton(text="Отмена", callback_data="upload_week:cancel")],
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def work_week_ids() -> tuple[str, str]:
    today = datetime.now().astimezone().date()
    current_monday = today - timedelta(days=today.isoweekday() - 1)
    current_year, current_week, _ = current_monday.isocalendar()
    previous_monday = current_monday - timedelta(weeks=1)
    previous_year, previous_week, _ = previous_monday.isocalendar()
    return (
        f"{current_year}-W{current_week:02d}",
        f"{previous_year}-W{previous_week:02d}",
    )


def work_week_picker() -> InlineKeyboardMarkup:
    current_week, previous_week = work_week_ids()
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📅 Текущая неделя",
                    callback_data=f"work_week:{current_week}",
                )
            ],
            [
                InlineKeyboardButton(
                    text="📅 Предыдущая неделя",
                    callback_data=f"work_week:{previous_week}",
                )
            ],
        ]
    )


def work_week_navigation() -> InlineKeyboardMarkup:
    current_week, previous_week = work_week_ids()
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="Текущая",
                    callback_data=f"work_week:{current_week}",
                ),
                InlineKeyboardButton(
                    text="Предыдущая",
                    callback_data=f"work_week:{previous_week}",
                ),
            ]
        ]
    )


def admin_panel_keyboard(admins: list[Any] | None = None) -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text="➕ Добавить администратора", callback_data="admin:add")],
        [InlineKeyboardButton(text="➖ Удалить администратора", callback_data="admin:remove")],
        [InlineKeyboardButton(text="📋 Список администраторов", callback_data="admin:list")],
        [InlineKeyboardButton(text="📊 Статус бота", callback_data="admin:status")],
        [InlineKeyboardButton(text="Закрыть", callback_data="admin:close")],
    ]
    if admins is not None:
        buttons = [
            [
                InlineKeyboardButton(
                    text=f"Удалить ID {admin['telegram_id']}",
                    callback_data=f"admin:remove_id:{admin['telegram_id']}",
                )
            ]
            for admin in admins
            if admin["source"] != "config"
        ] + [[InlineKeyboardButton(text="Назад", callback_data="admin:back")]]
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def is_admin(message: Message) -> bool:
    return (
        message.from_user is not None
        and message.chat.type == "private"
        and database.is_admin(message.from_user.id)
    )


def is_private(message: Message) -> bool:
    return message.chat.type == "private"


def is_private_admin_callback(callback: CallbackQuery) -> bool:
    return bool(
        callback.from_user
        and database.is_admin(callback.from_user.id)
        and callback.message
        and hasattr(callback.message, "chat")
        and callback.message.chat.type == "private"
    )


def log_action(message: Message, action: str) -> None:
    if message.from_user is not None:
        database.log_request(message.from_user.id, action)


async def discard_pending_upload(state: FSMContext, admin_id: int | None) -> None:
    data = await state.get_data()
    token = data.get("upload_token")
    if token and admin_id:
        database.delete_staged_payroll_upload(token, admin_id)


async def clear_workflow(state: FSMContext, telegram_id: int | None) -> None:
    await discard_pending_upload(state, telegram_id)
    await state.clear()


def roster_review(
    week_id: str, rows: list[tuple[str, str, list[str]]]
) -> tuple[list[str], list[str], list[str], list[tuple[str, str]]]:
    is_replacement = database.period_exists(week_id)
    comparison_week = week_id if is_replacement else adjacent_week_id(week_id, -1)
    if not database.period_exists(comparison_week):
        return [], [], [], []

    previous_display = database.period_employee_display_names(comparison_week)
    previous_names = set(previous_display)
    current_by_name = {normalized: display for normalized, display, _ in rows}
    current_values = {normalized: values for normalized, _, values in rows}
    incoming_names = set(current_by_name)
    added = sorted(incoming_names - previous_names, key=lambda name: current_by_name[name].casefold())
    removed = sorted(previous_names - incoming_names, key=lambda name: previous_display[name].casefold())
    changed: list[str] = []
    if is_replacement:
        previous_values = database.period_employee_values(comparison_week)
        changed = sorted(
            (
                current_by_name[name]
                for name in incoming_names & previous_names
                if previous_values[name] != current_values[name]
            ),
            key=str.casefold,
        )
    possible_renames: list[tuple[str, str]] = []
    remaining = set(added)
    for old_name in removed:
        if not remaining:
            break
        closest = max(
            remaining,
            key=lambda new_name: difflib.SequenceMatcher(
                None, old_name.replace(" ", ""), new_name.replace(" ", "")
            ).ratio(),
        )
        ratio = difflib.SequenceMatcher(
            None, old_name.replace(" ", ""), closest.replace(" ", "")
        ).ratio()
        if ratio >= 0.78:
            possible_renames.append((previous_display[old_name], current_by_name[closest]))
            remaining.remove(closest)
    return (
        [previous_display[name] for name in removed],
        [current_by_name[name] for name in added],
        changed,
        possible_renames,
    )


def work_review_changes(
    week_id: str,
    payroll_rows: list[tuple[str, str, list[str]]],
    work_rows: list[tuple[str, str, str, str]],
) -> list[str]:
    comparison_week = week_id if database.period_exists(week_id) else adjacent_week_id(week_id, -1)
    previous = database.period_work_rows(comparison_week)
    current: dict[str, list[dict[str, str]]] = {}
    for normalized, _, work_name, picks in work_rows:
        current.setdefault(normalized, []).append(
            {"work_name": work_name, "picks": picks}
        )
    return [
        display_name
        for normalized, display_name, _ in payroll_rows
        if previous.get(normalized, []) != current.get(normalized, [])
    ]


def payroll_diff_file(
    week_id: str,
    headers: list[str],
    rows: list[tuple[str, str, list[str]]],
    removed: list[str],
    added: list[str],
    changed: list[str],
    work_rows: list[tuple[str, str, str, str]] | None = None,
    changed_work: list[str] | None = None,
) -> BufferedInputFile:
    comparison_week = week_id if database.period_exists(week_id) else adjacent_week_id(week_id, -1)
    previous_headers = database.period_headers(comparison_week) or headers
    previous_values = database.period_employee_values(comparison_week)
    current_by_name = {
        normalized: (display, values) for normalized, display, values in rows
    }
    previous_work = database.period_work_rows(comparison_week)
    current_work: dict[str, list[dict[str, str]]] = {}
    for normalized, _, work_name, picks in work_rows or []:
        current_work.setdefault(normalized, []).append(
            {"work_name": work_name, "picks": picks}
        )
    names = list(dict.fromkeys(removed + added + changed + (changed_work or [])))
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(
        [
            "Изменение",
            "ФИО",
            "До (поля Excel)",
            "После (поля Excel)",
            "До (работа и пики)",
            "После (работа и пики)",
        ]
    )
    for name in names:
        normalized = normalize_name(name)
        if name in removed:
            change_label = "Удалён из списка"
        elif name in added:
            change_label = "Добавлен в список"
        elif name in changed and name in (changed_work or []):
            change_label = "Изменились расчёт и работа"
        elif name in changed:
            change_label = "Изменились данные строки"
        else:
            change_label = "Изменилась работа и количество пиков"
        writer.writerow(
            [
                change_label,
                name,
                json.dumps(dict(zip(previous_headers[1:], previous_values[normalized])), ensure_ascii=False)
                if normalized in previous_values
                else "",
                json.dumps(dict(zip(headers[1:], current_by_name[normalized][1])), ensure_ascii=False)
                if normalized in current_by_name
                else "",
                json.dumps(previous_work.get(normalized, []), ensure_ascii=False)
                if normalized in previous_work
                else "",
                json.dumps(current_work.get(normalized, []), ensure_ascii=False)
                if normalized in current_work
                else "",
            ]
        )
    return BufferedInputFile(
        output.getvalue().encode("utf-8-sig"),
        filename=f"payroll_review_{week_id}.csv",
    )


def format_amount(value: Any) -> str:
    if value is None or value == "—":
        return "—"
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return "—"
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    elif "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return str(value)
    if not amount.is_finite():
        return str(value)
    try:
        rounded = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return f"{rounded:,.2f}".replace(",", " ")
    except InvalidOperation:
        return str(value)


def format_quantity(value: Any) -> str:
    formatted = format_amount(value)
    return formatted[:-3] if formatted.endswith(".00") else formatted


def week_dates(week_id: str) -> list[date]:
    year, week_number = int(week_id[:4]), int(week_id[-2:])
    monday = date.fromisocalendar(year, week_number, 1)
    return [monday + timedelta(days=offset) for offset in range(7)]


def adjacent_week_id(week_id: str, offset: int) -> str:
    monday = week_dates(week_id)[0] + timedelta(weeks=offset)
    year, week_number, _ = monday.isocalendar()
    return f"{year}-W{week_number:02d}"


def week_label(week_id: str) -> str:
    monday, sunday = week_dates(week_id)[0], week_dates(week_id)[-1]
    if monday.year != sunday.year:
        return (
            f"{monday.day} {RU_MONTHS[monday.month - 1]} {monday.year} – "
            f"{sunday.day} {RU_MONTHS[sunday.month - 1]} {sunday.year}"
        )
    if monday.month != sunday.month:
        return (
            f"{monday.day} {RU_MONTHS[monday.month - 1]} – "
            f"{sunday.day} {RU_MONTHS[sunday.month - 1]} {sunday.year}"
        )
    return f"{monday.day}–{sunday.day} {RU_MONTHS[sunday.month - 1]} {sunday.year}"


def parse_week_id(value: str) -> str:
    if not WEEK_ID_PATTERN.fullmatch(value):
        raise ValueError("Формат недели должен быть ГГГГ-WНН, например 2026-W47.")
    year, week_number = int(value[:4]), int(value[-2:])
    try:
        date.fromisocalendar(year, week_number, 1)
    except ValueError as exc:
        raise ValueError("Указан несуществующий номер ISO-недели.") from exc
    return value


def cell_text(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip() or "—"


def validate_headers(values: tuple[Any, ...]) -> list[str]:
    headers = [cell_text(value) for value in values]
    while headers and headers[-1] == "—":
        headers.pop()
    if len(headers) not in {10, 11, 12}:
        raise ValueError(
            "Ожидаются ФИО, 7 дней, штрафы, необязательные доплаты/удержания и итог."
        )
    header_norms = [normalize_name(value).replace("ё", "е") for value in headers]
    if header_norms[0] not in {"фио", "ф.и.о.", "сотрудник", "employee", "name"}:
        raise ValueError("Первая колонка заголовка должна называться «ФИО».")
    if any(header == "—" for header in headers[1:8]):
        raise ValueError("Укажите заголовки всех семи колонок дней.")
    if "штраф" not in header_norms[8]:
        raise ValueError("Колонка после семи дней должна содержать заголовок «Штрафы».")
    if header_norms[-1] not in {"итого", "всего", "итог"}:
        raise ValueError("Последняя колонка должна называться «Итого».")
    adjustments = header_norms[9:-1]
    if len(adjustments) > 2:
        raise ValueError("Допускаются отдельные колонки доплат и удержаний.")
    if any(not any(word in header for word in ("доплат", "удерж")) for header in adjustments):
        raise ValueError("После штрафов допускаются только колонки доплат/удержаний.")
    return headers


def validate_xlsx_container(content: bytes) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            uncompressed_size = sum(info.file_size for info in archive.infolist())
    except zipfile.BadZipFile as exc:
        raise ValueError("Файл поврежден или не является корректным XLSX.") from exc
    if uncompressed_size > MAX_XLSX_UNCOMPRESSED_BYTES:
        raise ValueError("Содержимое Excel после распаковки превышает лимит 100 МБ.")


def parse_workbook_with_work(
    content: bytes,
) -> tuple[
    list[str],
    list[tuple[str, str, list[str]]],
    list[tuple[str, str, str, str]],
]:
    validate_xlsx_container(content)
    workbook = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    try:
        sheet = None
        header_error = None
        for candidate in workbook.worksheets:
            first_row = next(
                candidate.iter_rows(min_row=1, max_row=1, max_col=13, values_only=True),
                (),
            )
            if first_row and normalize_name(cell_text(first_row[0])).replace("ё", "е") in {
                "фио",
                "ф.и.о.",
                "сотрудник",
                "employee",
                "name",
            }:
                try:
                    validate_headers(first_row)
                except ValueError as exc:
                    header_error = exc
                    continue
                sheet = candidate
                break
        if sheet is None:
            if header_error:
                raise header_error
            raise ValueError("Не найдена вкладка с таблицей расчета и колонкой «ФИО».")
        iterator = sheet.iter_rows(values_only=True)
        try:
            headers = validate_headers(next(iterator))
        except StopIteration as exc:
            raise ValueError("Excel-файл пуст.") from exc

        rows: list[tuple[str, str, list[str]]] = []
        seen_names: set[str] = set()
        for row_number, row in enumerate(iterator, start=2):
            values = list(row[: len(headers)])
            if not values or all(value is None for value in values):
                continue
            values.extend([None] * (len(headers) - len(values)))
            display_name = cell_text(values[0]).strip()
            normalized_name = normalize_name(display_name)
            if not display_name or display_name == "—":
                raise ValueError(f"В строке {row_number} не указано ФИО.")
            if normalized_name in seen_names:
                raise ValueError(f"ФИО повторяется в строке {row_number}.")
            seen_names.add(normalized_name)
            rows.append(
                (normalized_name, display_name, [cell_text(value) for value in values[1:]])
            )
        if not rows:
            raise ValueError("В файле нет строк с сотрудниками.")
        work_rows = parse_work_sheet(workbook, seen_names)
        return headers, rows, work_rows
    finally:
        workbook.close()


def parse_workbook(content: bytes) -> tuple[list[str], list[tuple[str, str, list[str]]]]:
    headers, rows, _ = parse_workbook_with_work(content)
    return headers, rows


def parse_work_sheet(
    workbook: Any, payroll_employee_names: set[str]
) -> list[tuple[str, str, str, str]]:
    employee_headers = {"фио", "ф.и.о.", "сотрудник", "employee", "name"}
    work_sheet = None
    work_columns: tuple[int, int, int] | None = None
    for candidate in workbook.worksheets:
        first_row = next(
            candidate.iter_rows(min_row=1, max_row=1, max_col=8, values_only=True),
            (),
        )
        headers = [normalize_name(cell_text(value)).replace("ё", "е") for value in first_row]
        while headers and headers[-1] == "—":
            headers.pop()
        employee_index = next(
            (index for index, value in enumerate(headers) if value in employee_headers),
            None,
        )
        work_index = next(
            (
                index
                for index, value in enumerate(headers)
                if any(term in value for term in ("работ", "услуг", "service", "work"))
            ),
            None,
        )
        picks_index = next(
            (index for index, value in enumerate(headers) if "пик" in value or "pick" in value),
            None,
        )
        if employee_index is not None and work_index is not None and picks_index is not None:
            if work_sheet is not None:
                raise ValueError("В Excel найдено несколько вкладок с таблицей «Работа».")
            work_sheet = candidate
            work_columns = (employee_index, work_index, picks_index)
        elif normalize_name(candidate.title).replace("ё", "е") == "работа":
            raise ValueError(
                "На вкладке «Работа» нужны заголовки: «ФИО», «Работа», «Пики, шт.»."
            )

    if work_sheet is None or work_columns is None:
        return []

    employee_index, work_index, picks_index = work_columns
    work_rows: list[tuple[str, str, str, str]] = []
    for row_number, row in enumerate(
        work_sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        if not row or all(value is None for value in row):
            continue
        display_name = cell_text(row[employee_index] if employee_index < len(row) else None).strip()
        work_name = cell_text(row[work_index] if work_index < len(row) else None).strip()
        picks_value = row[picks_index] if picks_index < len(row) else None
        picks = cell_text(picks_value).strip()
        if not display_name or display_name == "—":
            raise ValueError(f"На вкладке «Работа», строка {row_number}: не указано ФИО.")
        if not work_name or work_name == "—":
            raise ValueError(
                f"На вкладке «Работа», строка {row_number}: не указано название работы."
            )
        if not picks or picks == "—":
            raise ValueError(
                f"На вкладке «Работа», строка {row_number}: не указано количество пиков."
            )
        try:
            count = Decimal(picks.replace(",", "."))
        except InvalidOperation as exc:
            raise ValueError(
                f"На вкладке «Работа», строка {row_number}: количество пиков должно быть числом."
            ) from exc
        if not count.is_finite() or count < 0:
            raise ValueError(
                f"На вкладке «Работа», строка {row_number}: количество пиков должно быть "
                "неотрицательным числом."
            )
        normalized_name = normalize_name(display_name)
        if normalized_name not in payroll_employee_names:
            raise ValueError(
                f"На вкладке «Работа», строка {row_number}: сотрудник «{display_name}» "
                "отсутствует в расчетной таблице."
            )
        work_rows.append((normalized_name, display_name, work_name, picks))
    return work_rows


def week_id_from_headers(headers: list[str]) -> str | None:
    day_headers = headers[1:8]
    parsed_dates: list[date] = []
    for header in day_headers:
        try:
            parsed_dates.append(date.fromisoformat(header))
        except ValueError:
            if any(_is_iso_date_header(value) for value in day_headers):
                raise ValueError(
                    "Не все семь заголовков дней содержат даты в формате ГГГГ-ММ-ДД."
                )
            return None
    if not parsed_dates:
        return None
    dates = parsed_dates
    if len(dates) != 7 or any(
        day != dates[0] + timedelta(days=index) for index, day in enumerate(dates)
    ):
        raise ValueError("Даты в заголовках должны идти подряд с понедельника по воскресенье.")
    iso_year, iso_week, weekday = dates[0].isocalendar()
    if weekday != 1 or dates[-1].isocalendar()[:2] != (iso_year, iso_week):
        raise ValueError("Даты в заголовках должны охватывать ровно одну неделю Пн–Вс.")
    return f"{iso_year}-W{iso_week:02d}"


def _is_iso_date_header(value: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value))


def render_payroll(record: dict[str, Any] | None, period_label: str) -> str:
    if record is None:
        return f"За {period_label} данных нет."
    if record["values"] is None:
        return f"За {period_label} расчетный лист не загружен."

    headers = record["headers"]
    values = record["values"]
    days = week_dates(record["week_id"])
    lines = [
        "📊 Расчётный лист",
        f"📅 Неделя: {week_label(record['week_id'])}",
        "",
        "━━━━━━━━━━━━━━",
        "",
    ]
    lines.extend(
        f"📆 {day:%d.%m} — {format_amount(value)}"
        for day, value in zip(days, values[:7])
    )
    lines.extend(["", "━━━━━━━━━━━━━━", ""])
    lines.append(f"⚠️ {headers[8]}: {format_amount(values[7])}")
    for label, value in zip(headers[9:-1], values[8:-1]):
        normalized = normalize_name(label)
        if "доплат" in normalized and "удерж" not in normalized:
            icon, caption = "➕", "Доплаты"
        elif "удерж" in normalized and "доплат" not in normalized:
            icon, caption = "➖", "Удержания"
        else:
            icon, caption = "➕➖", label
        lines.append(f"{icon} {caption}: {format_amount(value)}")
    lines.extend(["", f"💰 ИТОГО: {format_amount(values[-1])}"])
    return "\n".join(lines)


def render_work(
    week_id: str, rows: list[dict[str, str]] | None
) -> str:
    if not rows:
        return (
            f"📦 За неделю {week_label(week_id)} данные о работе и пиках не загружены."
        )
    lines = [
        "📦 Работа и пики",
        f"📅 Неделя: {week_label(week_id)}",
        "",
        "━━━━━━━━━━━━━━",
        "",
    ]
    total = Decimal("0")
    for row in rows:
        amount = Decimal(row["picks"].replace(",", "."))
        total += amount
        lines.append(f"• {row['work_name']} — {format_quantity(row['picks'])} шт.")
    lines.extend(["", "━━━━━━━━━━━━━━", "", f"Итого пиков: {format_quantity(total)} шт."])
    return "\n".join(lines)


@router.message(Command("start"))
async def start(message: Message, state: FSMContext) -> None:
    await discard_pending_upload(state, message.from_user.id if message.from_user else None)
    await state.clear()
    log_action(message, "start")
    if not is_private(message):
        await message.answer("Откройте личный чат с ботом и отправьте /start.")
        return
    owner = is_admin(message)
    if message.from_user and _is_linked_telegram_id(message.from_user.id):
        await message.answer(
            "Вы вошли. Выберите расчетный лист:", reply_markup=keyboard(owner)
        )
        return
    await message.answer(
        "Здравствуйте! Нажмите «🔐 Войти» и укажите ФИО и пароль."
        if not owner
        else "Меню администратора открыто. Для личного расчета нажмите «🔐 Войти».",
        reply_markup=keyboard(owner),
    )


def _is_linked_telegram_id(telegram_id: int) -> bool:
    return database.is_telegram_linked(telegram_id)


@router.message(Command("login"))
async def login_start(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    if not is_private(message):
        await message.answer("Для безопасности вход доступен только в личном чате с ботом.")
        return
    await state.set_state(Login.name)
    log_action(message, "login_started")
    await message.answer("Введите ФИО так, как оно указано в расчетном листе.")


@router.message(F.text.in_({"🔐 Войти", "Войти"}))
async def login_button(message: Message, state: FSMContext) -> None:
    await login_start(message, state)


@router.message(Login.name, F.text, ~F.text.startswith("/"), ~F.text.in_(MENU_BUTTONS))
async def login_name(message: Message, state: FSMContext) -> None:
    name = (message.text or "").strip()
    if not name or len(name) > 200:
        await message.answer("Введите корректное ФИО.")
        return
    await state.update_data(employee_name=name)
    await state.set_state(Login.password)
    await message.answer("Теперь введите пароль.")


@router.message(Login.password, F.text, ~F.text.startswith("/"), ~F.text.in_(MENU_BUTTONS))
async def login_password(message: Message, state: FSMContext) -> None:
    telegram_id = message.from_user.id
    now = time.time()
    locked_until = database.login_lockout(telegram_id, now)
    if locked_until > now:
        await state.clear()
        await message.answer("Слишком много попыток. Попробуйте через 15 минут.")
        log_action(message, "login_rate_limited")
        return

    data = await state.get_data()
    employee = database.employee_by_name(normalize_name(data.get("employee_name", "")))
    password = message.text or ""
    valid_password = False
    if employee and employee["password_hash"] and len(password.encode("utf-8")) <= 72:
        valid_password = await asyncio.to_thread(
            bcrypt.checkpw, password.encode("utf-8"), bytes(employee["password_hash"])
        )
    await state.clear()
    if not valid_password:
        locked_until = database.record_login_failure(
            telegram_id, now, MAX_LOGIN_FAILURES, LOGIN_WINDOW_SECONDS, LOGIN_LOCK_SECONDS
        )
        log_action(message, "login_failed")
        if locked_until > now:
            await message.answer("Слишком много попыток. Попробуйте через 15 минут.")
        else:
            await message.answer("ФИО или пароль неверны. Попробуйте снова: /login")
        return

    if not database.link_telegram_account(employee["id"], telegram_id):
        await message.answer(
            "Эта учетная запись уже связана с другим Telegram-аккаунтом. "
            "Обратитесь к владельцу бота."
        )
        log_action(message, "login_account_already_linked")
        return
    database.clear_login_failures(telegram_id)
    log_action(message, "login_succeeded")
    await message.answer(
        "Вход выполнен. Выберите расчетный лист:",
        reply_markup=keyboard(database.is_admin(telegram_id)),
    )


@router.message(Login.name, ~F.text)
async def login_name_non_text(message: Message) -> None:
    await message.answer("Отправьте ФИО текстом или /cancel для отмены.")


@router.message(Login.password, ~F.text)
async def login_password_non_text(message: Message) -> None:
    await message.answer("Отправьте пароль текстом или /cancel для отмены.")


@router.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    owner = is_admin(message)
    await message.answer("Действие отменено. Главное меню:", reply_markup=keyboard(owner))


@router.message(Command("logout"))
@router.message(F.text.in_({"🚪 Выйти", "Выйти"}))
async def logout(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    if not is_private(message):
        await message.answer("Выход доступен только в личном чате с ботом.")
        return
    if message.from_user:
        database.unlink_telegram_id(message.from_user.id)
    log_action(message, "logout")
    await message.answer(
        "Вы вышли из аккаунта. Для повторного входа нажмите «🔐 Войти».",
        reply_markup=keyboard(is_admin(message)),
    )


@router.message(
    F.text.in_(
        {
            "📅 Текущая неделя",
            "📅 Предыдущая неделя",
            "Текущая неделя",
            "Предыдущая неделя",
        }
    )
)
async def show_payroll(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    if not is_private(message):
        await message.answer("Расчётный лист можно посмотреть только в личном чате с ботом.")
        return
    if not message.from_user or not _is_linked_telegram_id(message.from_user.id):
        log_action(message, "payroll_unauthorized")
        await message.answer("Сначала войдите кнопкой «🔐 Войти».")
        return
    index = 0 if "Текущая неделя" in (message.text or "") else 1
    today = datetime.now().astimezone().date()
    monday = today - timedelta(days=today.isoweekday() - 1)
    selected_monday = monday - timedelta(weeks=index)
    iso_year, iso_week, _ = selected_monday.isocalendar()
    week_id = f"{iso_year}-W{iso_week:02d}"
    record = database.payroll_for_week(message.from_user.id, week_id)
    period_label = "текущую неделю" if index == 0 else "предыдущую неделю"
    log_action(message, "payroll_current" if index == 0 else "payroll_previous")
    if record is None:
        await message.answer(
            f"За {period_label} ({week_label(week_id)}) данные ещё не загружены.",
            reply_markup=keyboard(is_admin(message)),
        )
        return
    if record["values"] is None:
        await message.answer(
            f"В расчёте за {week_label(week_id)} ваши данные не найдены. "
            "Обратитесь к администратору.",
            reply_markup=keyboard(is_admin(message)),
        )
        return
    await message.answer(
        render_payroll(record, period_label),
        reply_markup=keyboard(is_admin(message)),
    )


@router.message(F.text == "📦 Работа и пики")
async def show_work_picker(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    if not is_private(message):
        await message.answer("Работу и пики можно посмотреть только в личном чате с ботом.")
        return
    if not message.from_user or not _is_linked_telegram_id(message.from_user.id):
        log_action(message, "work_unauthorized")
        await message.answer("Сначала войдите кнопкой «🔐 Войти».")
        return
    log_action(message, "work_week_picker")
    await message.answer("Выберите неделю для просмотра:", reply_markup=work_week_picker())


@router.callback_query(F.data.startswith("work_week:"))
async def show_work_week(callback: CallbackQuery) -> None:
    if (
        not callback.from_user
        or not callback.message
        or not hasattr(callback.message, "chat")
        or callback.message.chat.type != "private"
    ):
        await callback.answer("Просмотр доступен только в личном чате.", show_alert=True)
        return
    if not _is_linked_telegram_id(callback.from_user.id):
        await callback.answer("Сначала войдите в аккаунт.", show_alert=True)
        return
    try:
        week_id = parse_week_id((callback.data or "").partition(":")[2])
    except ValueError:
        await callback.answer("Неделя указана неверно.", show_alert=True)
        return

    if not database.period_exists(week_id):
        text = f"За неделю {week_label(week_id)} расчёт ещё не загружен."
    else:
        rows = database.work_rows_for_employee(callback.from_user.id, week_id)
        if rows is None:
            await callback.answer("Сначала войдите в аккаунт.", show_alert=True)
            return
        text = (
            render_work(week_id, rows)
            if rows
            else f"За неделю {week_label(week_id)} данные о работе и пиках в Excel отсутствуют."
        )
    await callback.answer()
    if callback.message.text != text:
        await callback.message.edit_text(text, reply_markup=work_week_navigation())
    database.log_request(callback.from_user.id, f"work_view:{week_id}")


@router.message(Command("upload"))
async def upload_start(message: Message, command: CommandObject, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    await discard_pending_upload(state, message.from_user.id)
    await state.clear()
    if not command.args:
        await state.set_state(Upload.week)
        log_action(message, "payroll_upload_week_picker")
        await message.answer("Выберите неделю для загрузки:", reply_markup=week_picker())
        return
    await begin_upload(message, state, command.args.strip())


@router.message(F.text.in_({"📤 Загрузить Excel", "Загрузить Excel"}))
async def upload_button(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    await discard_pending_upload(state, message.from_user.id)
    await state.clear()
    await state.set_state(Upload.week)
    log_action(message, "payroll_upload_week_picker")
    await message.answer("Выберите неделю для загрузки:", reply_markup=week_picker())


async def begin_upload(message: Message, state: FSMContext, week_value: str) -> None:
    try:
        week_id = parse_week_id(week_value)
    except ValueError as exc:
        log_action(message, "payroll_upload_invalid_week")
        await message.answer(f"{exc}\nПример: /upload 2026-W47")
        return
    await state.update_data(
        week_id=week_id,
        expected_uploaded_at=database.period_uploaded_at(week_id),
    )
    if database.period_exists(week_id):
        await state.set_state(Upload.confirm_replace)
        await message.answer(
            f"За {week_id} уже есть загруженный расчёт. Заменить данные этой недели?",
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[
                    [
                        InlineKeyboardButton(
                            text="Заменить",
                            callback_data=f"upload_replace:{week_id}",
                        ),
                        InlineKeyboardButton(
                            text="Отмена",
                            callback_data="upload_replace:cancel",
                        ),
                    ]
                ]
            ),
        )
        return
    await state.update_data(week_id=week_id)
    await state.set_state(Upload.document)
    log_action(message, f"payroll_upload_started:{week_id}")
    await message.answer(
        f"Пришлите XLSX-файл для недели {week_id}. Для отмены отправьте /cancel."
    )


@router.callback_query(F.data.startswith("upload_week:"))
async def upload_week_selected(callback: CallbackQuery, state: FSMContext) -> None:
    if not is_private_admin_callback(callback):
        await callback.answer("Только владелец может загружать расчет.", show_alert=True)
        return
    selection = (callback.data or "").partition(":")[2]
    await callback.answer()
    if selection == "cancel":
        await discard_pending_upload(state, callback.from_user.id)
        await state.clear()
        if callback.message:
            await callback.message.edit_text("Загрузка отменена.")
        return
    if selection == "custom":
        await discard_pending_upload(state, callback.from_user.id)
        await state.set_state(Upload.week)
        if callback.message:
            await callback.message.edit_text(
                "Введите неделю в формате ГГГГ-WНН, например 2026-W47. Для отмены /cancel."
            )
        return

    try:
        week_id = parse_week_id(selection)
    except ValueError as exc:
        await state.clear()
        if callback.message:
            await callback.message.answer(f"Не удалось выбрать неделю: {exc}")
        return
    if callback.message:
        await begin_upload(callback.message, state, week_id)


@router.callback_query(F.data.startswith("upload_replace:"))
async def upload_replace_selected(callback: CallbackQuery, state: FSMContext) -> None:
    if not is_private_admin_callback(callback):
        await callback.answer("Только администратор может загружать расчёты.", show_alert=True)
        return
    selection = (callback.data or "").partition(":")[2]
    await callback.answer()
    if selection == "cancel":
        await state.clear()
        if callback.message:
            await callback.message.edit_text("Загрузка отменена.")
        return
    try:
        week_id = parse_week_id(selection)
    except ValueError as exc:
        await state.clear()
        if callback.message:
            await callback.message.answer(f"Не удалось подтвердить замену: {exc}")
        return
    if not database.period_exists(week_id):
        await state.clear()
        if callback.message:
            await callback.message.answer(
                f"Расчёт {week_id} уже отсутствует. Начните загрузку заново."
            )
        return
    data = await state.get_data()
    if data.get("expected_uploaded_at") != database.period_uploaded_at(week_id):
        await state.clear()
        if callback.message:
            await callback.message.answer(
                "Эту неделю уже изменил другой администратор. Начните загрузку заново."
            )
        return
    await state.update_data(week_id=week_id)
    await state.set_state(Upload.document)
    if callback.message:
        await callback.message.edit_text(
            f"Заменяем данные {week_id}. Теперь отправьте XLSX-файл или /cancel."
        )


@router.message(Upload.week, F.text, ~F.text.startswith("/"), ~F.text.in_(MENU_BUTTONS))
async def upload_custom_week(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        await state.clear()
        await message.answer("Команда доступна только владельцу бота.")
        return
    await begin_upload(message, state, (message.text or "").strip())


@router.message(Upload.week, ~F.text)
async def upload_week_non_text(message: Message) -> None:
    await message.answer("Выберите неделю кнопкой или отправьте её текстом, например 2026-W47.")


@router.message(Upload.document, F.document)
async def upload_document(message: Message, state: FSMContext) -> None:
    if not is_admin(message) or not message.document:
        await state.clear()
        return
    document = message.document
    if not document.file_name or not document.file_name.lower().endswith(".xlsx"):
        log_action(message, "payroll_upload_invalid_file_type")
        await message.answer("Нужен файл формата .xlsx. Попробуйте еще раз или /cancel.")
        return
    if document.file_size is not None and document.file_size > MAX_XLSX_BYTES:
        log_action(message, "payroll_upload_too_large")
        await message.answer("Файл больше допустимых 10 МБ.")
        return

    content = io.BytesIO()
    try:
        await message.bot.download(document, destination=content)
    except TelegramAPIError:
        logger.exception("Could not download an XLSX upload from admin %s", message.from_user.id)
        await message.answer("Не удалось скачать файл из Telegram. Попробуйте отправить его ещё раз.")
        return
    if content.tell() > MAX_XLSX_BYTES:
        log_action(message, "payroll_upload_too_large")
        await message.answer("Файл больше допустимых 10 МБ.")
        return
    try:
        headers, rows, work_rows = parse_workbook_with_work(content.getvalue())
        data = await state.get_data()
        week_id = data["week_id"]
        workbook_week_id = week_id_from_headers(headers)
        if workbook_week_id and workbook_week_id != week_id:
            raise ValueError(
                f"Даты Excel относятся к {workbook_week_id}, "
                f"но выбрана неделя {week_id}."
            )
        expected_uploaded_at = data.get("expected_uploaded_at")
        if database.period_uploaded_at(week_id) != expected_uploaded_at:
            raise database.ConcurrentPayrollUpdateError(
                "Эту неделю уже изменил другой администратор. Загрузка отменена; проверьте данные."
            )
        removed, added, changed, possible_renames = roster_review(week_id, rows)
        changed_work = work_review_changes(week_id, rows, work_rows)
        if possible_renames:
            details = "\n".join(
                f"• {old_name} → {new_name}" for old_name, new_name in possible_renames[:10]
            )
            await state.clear()
            logger.warning(
                "Payroll upload roster has suspicious name changes: week=%s admin=%s count=%d",
                week_id,
                message.from_user.id,
                len(possible_renames),
            )
            log_action(message, "payroll_upload_suspicious_roster")
            await message.answer(
                "В списке есть ФИО, похожие на переименованные записи. "
                "Чтобы не выдать чужой расчёт, загрузка остановлена. "
                "Сверьте ФИО с исходным Excel и исправьте возможные опечатки:\n"
                f"{details}"
            )
            return
        token = secrets.token_hex(16)
        database.stage_payroll_upload(
            token,
            message.from_user.id,
            week_id,
            headers,
            rows,
            expected_uploaded_at,
            time.time(),
            work_rows=work_rows,
        )
        await state.update_data(upload_token=token)
        if removed or added or changed or changed_work:
            await state.set_state(Upload.review)
            review_file = payroll_diff_file(
                week_id, headers, rows, removed, added, changed, work_rows, changed_work
            )
            await message.answer_document(
                review_file,
                caption=(
                    f"Изменения для {week_id}: добавлено {len(added)}, "
                    f"удалено {len(removed)}, расчётных строк изменено {len(changed)}, "
                    f"строк работы изменено {len(changed_work)}. "
                    "Сверьте CSV. Старые данные пока не менялись."
                ),
                reply_markup=InlineKeyboardMarkup(
                    inline_keyboard=[
                        [
                            InlineKeyboardButton(
                                text="Подтвердить загрузку",
                                callback_data=f"upload_review:confirm:{token}",
                            ),
                            InlineKeyboardButton(
                                text="Отмена",
                                callback_data=f"upload_review:cancel:{token}",
                            ),
                        ]
                    ]
                ),
            )
            return
        await commit_staged_upload(message, state, token, message.from_user.id)
    except database.ConcurrentPayrollUpdateError as exc:
        await state.clear()
        logger.warning("Concurrent payroll upload rejected for week %s: %s", week_id, exc)
        log_action(message, "payroll_upload_conflict")
        await message.answer(str(exc))
        return
    except (
        ValueError,
        InvalidFileException,
        zipfile.BadZipFile,
        OSError,
        KeyError,
        ElementTree.ParseError,
    ) as exc:
        logger.warning(
            "Rejected payroll upload from admin %s: %s",
            message.from_user.id,
            exc,
        )
        log_action(message, "payroll_upload_rejected")
        await message.answer(f"Файл не загружен: {exc}")
        return
    except sqlite3.Error:
        logger.exception("Database error while importing payroll week %s", week_id)
        await message.answer("Не удалось сохранить расчёт из-за ошибки базы данных. Попробуйте позже.")
        return


async def commit_staged_upload(
    message: Message, state: FSMContext, token: str, admin_id: int
) -> None:
    if not database.is_admin(admin_id):
        await state.clear()
        await message.answer("Доступ к загрузке отозван. Загрузка отменена.")
        return
    staged = database.get_staged_payroll_upload(token, admin_id, time.time())
    if staged is None:
        await state.clear()
        await message.answer("Срок подтверждения истёк. Начните загрузку заново.")
        return
    week_id = staged["week_id"]
    try:
        database.save_payroll(
            week_id,
            staged["headers"],
            [tuple(row) for row in staged["rows"]],
            expected_uploaded_at=staged["expected_uploaded_at"],
            check_version=True,
            work_rows=[tuple(row) for row in staged["work_rows"]],
        )
    except database.ConcurrentPayrollUpdateError as exc:
        logger.warning("Concurrent payroll upload rejected for week %s: %s", week_id, exc)
        database.delete_staged_payroll_upload(token, admin_id)
        await state.clear()
        log_action_from_id(admin_id, "payroll_upload_conflict")
        await message.answer(str(exc), reply_markup=keyboard(True))
        return
    except sqlite3.Error:
        logger.exception("Database error while committing payroll week %s", week_id)
        await message.answer(
            "Не удалось сохранить расчёт из-за ошибки базы данных. Попробуйте позже."
        )
        return
    database.delete_staged_payroll_upload(token, admin_id)
    await state.clear()
    log_action_from_id(admin_id, f"payroll_uploaded:{week_id}")
    await message.answer(
        f"✅ Расчёт за {week_id} загружен: {len(staged['rows'])} сотрудников. "
        "Другие недели, пароли и привязки сохранены.",
        reply_markup=keyboard(True),
    )


@router.callback_query(F.data.startswith("upload_review:"))
async def upload_review_selected(callback: CallbackQuery, state: FSMContext) -> None:
    if not is_private_admin_callback(callback):
        await callback.answer("Доступ только для администратора.", show_alert=True)
        return
    if not callback.message or len((callback.data or "").split(":")) != 3:
        await callback.answer("Кнопка подтверждения недействительна.", show_alert=True)
        return
    _, action, token = (callback.data or "").split(":")
    if len(token) != 32 or not re.fullmatch(r"[a-f0-9]{32}", token):
        await callback.answer("Кнопка подтверждения недействительна.", show_alert=True)
        return
    staged = database.get_staged_payroll_upload(token, callback.from_user.id, time.time())
    if staged is None:
        await state.clear()
        await callback.answer("Срок подтверждения истёк.", show_alert=True)
        await callback.message.edit_text("Срок подтверждения истёк. Начните загрузку заново.")
        return
    if action == "cancel":
        database.delete_staged_payroll_upload(token, callback.from_user.id)
        await state.clear()
        await callback.answer()
        await callback.message.edit_text("Загрузка отменена; данные в базе не менялись.")
        return
    if action != "confirm":
        await callback.answer("Действие недействительно.", show_alert=True)
        return
    await callback.answer()
    await commit_staged_upload(callback.message, state, token, callback.from_user.id)


@router.message(Upload.document, F.text, ~F.text.startswith("/"), ~F.text.in_(MENU_BUTTONS))
async def upload_non_document(message: Message) -> None:
    await message.answer("Пришлите Excel-файл .xlsx или отправьте /cancel.")


@router.message(Upload.document, ~F.document, ~F.text)
async def upload_non_document_media(message: Message) -> None:
    await message.answer("Пришлите файл .xlsx или отправьте /cancel.")


@router.message(F.text.in_({"👥 Сотрудники", "Сотрудники"}))
async def employee_list(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "admin_access_denied")
        await message.answer("Эта функция доступна только администраторам.")
        return
    await clear_workflow(state, message.from_user.id)
    employees = database.employee_directory()
    if not employees:
        await message.answer("Список сотрудников пока пуст.", reply_markup=keyboard(True))
        return
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(["ФИО", "Пароль настроен", "Telegram привязан"])
    writer.writerows(
        [
            [
                employee["display_name"],
                "Да" if employee["has_password"] else "Нет",
                "Да" if employee["is_linked"] else "Нет",
            ]
            for employee in employees
        ]
    )
    log_action(message, "employee_directory_requested")
    await message.answer_document(
        BufferedInputFile(
            output.getvalue().encode("utf-8-sig"), filename="employees.csv"
        ),
        caption=f"Сотрудников: {len(employees)}. Файл содержит только список и статус входа.",
        reply_markup=keyboard(True),
    )


def make_password() -> str:
    return secrets.token_urlsafe(12)


@router.message(Command("generate_passwords"))
@router.message(F.text.in_({"🔑 Пароли сотрудников", "Пароли сотрудникам"}))
async def generate_passwords(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    await clear_workflow(state, message.from_user.id)
    employees = database.employee_names_without_passwords()
    if not employees:
        await message.answer(
            "Нет новых сотрудников без пароля. Пароли ранее созданных сотрудников не показываются повторно."
        )
        return

    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(["ФИО", "Пароль"])
    generated: set[str] = set()
    password_hashes: list[tuple[int, bytes]] = []
    for employee in employees:
        password = make_password()
        while password in generated:
            password = make_password()
        generated.add(password)
        password_hash = await asyncio.to_thread(
            bcrypt.hashpw, password.encode("utf-8"), bcrypt.gensalt(rounds=12)
        )
        password_hashes.append((employee["id"], password_hash))
        writer.writerow([employee["display_name"], password])

    database.set_passwords(password_hashes)
    file = BufferedInputFile(
        output.getvalue().encode("utf-8-sig"), filename="employee_passwords.csv"
    )
    log_action(message, "passwords_generated")
    try:
        await message.answer_document(
            file,
            caption=(
                f"Созданы пароли для {len(employees)} сотрудников. "
                "Скачайте файл и передайте каждому его пароль безопасным способом."
            ),
        )
    except TelegramAPIError:
        database.clear_passwords_if_match(password_hashes)
        logger.exception("Could not deliver employee password file to admin %s", message.from_user.id)
        raise


@router.message(Command("reset_password"))
async def reset_password_command(
    message: Message, command: CommandObject, state: FSMContext
) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    name = (command.args or "").strip()
    await discard_pending_upload(state, message.from_user.id)
    if not name:
        await state.clear()
        await state.set_state(ResetPassword.employee_name)
        await message.answer("Введите ФИО сотрудника для сброса пароля.")
        return
    await state.clear()
    await issue_password(message, name)


@router.message(F.text.in_({"🔄 Сброс пароля", "Сбросить пароль"}))
async def reset_password_button(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    await clear_workflow(state, message.from_user.id)
    await state.set_state(ResetPassword.employee_name)
    await message.answer("Введите ФИО сотрудника для сброса пароля.")


@router.message(
    ResetPassword.employee_name, F.text, ~F.text.startswith("/"), ~F.text.in_(MENU_BUTTONS)
)
async def reset_password_name(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        await state.clear()
        await message.answer("Команда доступна только владельцу бота.")
        return
    await state.clear()
    await issue_password(message, (message.text or "").strip())


@router.message(ResetPassword.employee_name, ~F.text)
async def reset_password_non_text(message: Message) -> None:
    await message.answer("Отправьте ФИО сотрудника текстом или /cancel для отмены.")


async def issue_password(message: Message, name: str) -> None:
    employee = database.employee_by_name(normalize_name(name)) if name else None
    if employee is None:
        await message.answer(
            "Сотрудник не найден. Проверьте ФИО в загруженном Excel и попробуйте снова."
        )
        return

    password = make_password()
    password_hash = await asyncio.to_thread(
        bcrypt.hashpw, password.encode("utf-8"), bcrypt.gensalt(rounds=12)
    )
    previous_hash = bytes(employee["password_hash"]) if employee["password_hash"] else None
    database.set_password(employee["id"], password_hash)
    try:
        await message.answer(f"Новый пароль для сотрудника:\n{password}")
    except TelegramAPIError:
        database.restore_password(employee["id"], previous_hash)
        logger.exception("Could not deliver a reset password to admin %s", message.from_user.id)
        raise
    database.unlink_telegram_employee(employee["id"])
    log_action(message, "password_reset")


@router.message(Command("unbind"))
async def unbind_employee(
    message: Message, command: CommandObject, state: FSMContext
) -> None:
    if not is_admin(message):
        log_action(message, "owner_access_denied")
        await message.answer("Команда доступна только владельцу бота.")
        return
    await clear_workflow(state, message.from_user.id)
    name = (command.args or "").strip()
    if not name:
        await message.answer("Укажите ФИО: /unbind Иванов Иван Иванович")
        return
    found = database.unlink_telegram_account(normalize_name(name))
    log_action(message, "employee_unbound")
    await message.answer(
        "Привязка снята. Сотрудник сможет войти заново."
        if found
        else "Сотрудник с таким ФИО не найден."
    )


@router.message(F.text.in_({"⚙️ Администрирование", "Администрирование"}))
async def admin_management(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "admin_access_denied")
        await message.answer("Эта функция доступна только администраторам.")
        return
    await clear_workflow(state, message.from_user.id)
    log_action(message, "admin_panel_opened")
    await message.answer(
        "Управление ботом. Владельцы из .env защищены от удаления.",
        reply_markup=admin_panel_keyboard(),
    )


@router.callback_query(F.data.startswith("admin:"))
async def admin_management_callback(callback: CallbackQuery, state: FSMContext) -> None:
    if not is_private_admin_callback(callback):
        await callback.answer("Доступ только для администратора.", show_alert=True)
        return
    if not hasattr(callback.message, "edit_text"):
        await callback.answer("Сообщение панели недоступно.", show_alert=True)
        return

    action = (callback.data or "").split(":")
    operation = action[1] if len(action) > 1 else ""
    await discard_pending_upload(state, callback.from_user.id)
    await state.clear()
    await callback.answer()
    if operation == "close":
        await state.clear()
        await callback.message.edit_text("Панель закрыта. Команды доступны в главном меню.")
    elif operation == "back":
        await callback.message.edit_text(
            "Управление ботом. Владельцы из .env защищены от удаления.",
            reply_markup=admin_panel_keyboard(),
        )
    elif operation == "list":
        admins = database.list_admins()
        lines = ["Администраторы:"]
        lines.extend(
            f"• {admin['telegram_id']}"
            + (" — из конфигурации" if admin["source"] == "config" else "")
            for admin in admins
        )
        await callback.message.edit_text(
            "\n".join(lines), reply_markup=admin_panel_keyboard()
        )
    elif operation == "status":
        await callback.message.edit_text(
            bot_status_text(), reply_markup=admin_panel_keyboard()
        )
    elif operation == "add":
        await state.set_state(AdminManagement.add_admin_id)
        await callback.message.edit_text(
            "Отправьте числовой Telegram ID нового администратора. Для отмены — /cancel."
        )
    elif operation == "remove":
        admins = database.list_admins()
        removable = [admin for admin in admins if admin["source"] != "config"]
        if not removable:
            await callback.message.edit_text(
                "Нет администраторов, которых можно удалить.",
                reply_markup=admin_panel_keyboard(),
            )
        else:
            await callback.message.edit_text(
                "Выберите администратора для удаления:",
                reply_markup=admin_panel_keyboard(admins),
            )
    elif operation == "remove_id" and len(action) == 3 and action[2].isdigit():
        target_id = int(action[2])
        if target_id == callback.from_user.id:
            await callback.message.edit_text(
                "Нельзя удалить собственную учетную запись.",
                reply_markup=admin_panel_keyboard(),
            )
            return
        if target_id in CONFIG_ADMIN_IDS:
            await callback.message.edit_text(
                "Администратор из конфигурации защищён от удаления.",
                reply_markup=admin_panel_keyboard(),
            )
            return
        await callback.message.edit_text(
            f"Подтвердите удаление администратора {target_id}.",
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[
                    [
                        InlineKeyboardButton(
                            text="Удалить",
                            callback_data=f"admin:confirm_remove:{target_id}",
                        ),
                        InlineKeyboardButton(text="Отмена", callback_data="admin:back"),
                    ]
                ]
            ),
        )
    elif operation == "confirm_remove" and len(action) == 3 and action[2].isdigit():
        target_id = int(action[2])
        result = database.remove_admin(target_id, callback.from_user.id)
        messages = {
            "removed": f"Администратор {target_id} удалён.",
            "not_found": "Этот администратор уже удалён.",
            "protected": "Администратор из конфигурации защищён от удаления.",
            "self": "Нельзя удалить собственную учетную запись.",
            "last_admin": "Нельзя удалить последнего администратора.",
        }
        log_action_from_id(
            callback.from_user.id,
            "admin_removed" if result == "removed" else "admin_remove_rejected",
        )
        await callback.message.edit_text(messages[result], reply_markup=admin_panel_keyboard())


@router.message(
    AdminManagement.add_admin_id,
    F.text,
    ~F.text.startswith("/"),
    ~F.text.in_(MENU_BUTTONS),
)
async def add_admin_id(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        await state.clear()
        await message.answer("Эта функция доступна только администраторам.")
        return
    value = (message.text or "").strip()
    if not value.isascii() or not value.isdigit() or int(value) <= 0:
        await message.answer("Отправьте положительный числовой Telegram ID или /cancel.")
        return
    target_id = int(value)
    added = database.add_admin(target_id, message.from_user.id)
    await state.clear()
    log_action(message, "admin_added" if added else "admin_add_duplicate")
    await message.answer(
        f"Администратор {target_id} добавлен."
        if added
        else f"Аккаунт {target_id} уже является администратором.",
        reply_markup=keyboard(True),
    )


@router.message(AdminManagement.add_admin_id, ~F.text)
async def add_admin_non_text(message: Message) -> None:
    await message.answer("Отправьте числовой Telegram ID текстом или /cancel для отмены.")


@router.message(Command("status"))
async def bot_status(message: Message, state: FSMContext) -> None:
    if not is_admin(message):
        log_action(message, "admin_access_denied")
        await message.answer("Эта функция доступна только администраторам.")
        return
    await clear_workflow(state, message.from_user.id)
    await message.answer(bot_status_text(), reply_markup=keyboard(True))
    log_action(message, "bot_status_requested")


def bot_status_text() -> str:
    uptime_seconds = int(time.monotonic() - STARTED_AT)
    hours, remainder = divmod(uptime_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    latest = database.latest_period() or "нет"
    return (
        "📊 Статус бота: работает\n"
        f"⏱ Время с запуска: {hours} ч {minutes} мин {seconds} сек\n"
        f"👥 Сотрудников: {database.employee_count()}\n"
        f"🔗 Привязано аккаунтов: {database.linked_employee_count()}\n"
        f"📅 Последняя загруженная неделя: {latest}"
    )


def log_action_from_id(telegram_id: int, action: str) -> None:
    database.log_request(telegram_id, action)


@router.message(Command("help"))
async def help_command(message: Message, state: FSMContext) -> None:
    await clear_workflow(state, message.from_user.id if message.from_user else None)
    log_action(message, "help")
    text = (
        "Сотрудник: кнопки «Войти», «Текущая неделя», «Предыдущая неделя», «Выйти».\n"
        "Администратор: кнопки меню — загрузка Excel, сотрудники, пароли и управление. "
        "Также доступны /upload ГГГГ-WНН и /status; "
        "/generate_passwords — получить CSV паролей новых сотрудников; "
        "/reset_password ФИО — выдать новый пароль и сбросить активную привязку; "
        "/unbind ФИО — разрешить повторную привязку аккаунта."
    )
    await message.answer(text)


@dispatcher.errors()
async def unhandled_update_error(event: ErrorEvent) -> bool:
    logger.error(
        "Unhandled error while processing Telegram update",
        exc_info=(type(event.exception), event.exception, event.exception.__traceback__),
    )
    update = event.update
    try:
        if update.callback_query:
            callback = update.callback_query
            if callback.message and callback.message.chat.type == "private":
                await callback.answer("Не удалось выполнить действие. Попробуйте ещё раз.")
        elif update.message and update.message.chat.type == "private":
            await update.message.answer(
                "Не удалось выполнить действие из-за внутренней ошибки. Попробуйте позже."
            )
    except TelegramAPIError:
        logger.exception("Could not send an error notification to the user")
    return True


async def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    os.umask(0o077)
    database.initialize(CONFIG_ADMIN_IDS)
    bot = Bot(token=BOT_TOKEN)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
