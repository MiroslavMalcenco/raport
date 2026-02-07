#!/usr/bin/env python3
"""merge_pallets.py

Скрипт для обработки файлов Pallet 1.xlsx .. Pallet 30.xlsx, объединения,
очистки данных и обогащения по спецификации.

Требования: pandas, openpyxl
"""
from pathlib import Path
import logging
import sys
import os
import subprocess
from typing import List, Dict, Optional

import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# При запуске как --noconsole exe перенаправляем stderr в файл,
# чтобы ошибки не пропадали бесследно.
if getattr(sys, 'frozen', False):
    _crash_log = Path(sys.executable).parent / 'crash.log'
    try:
        sys.stderr = open(str(_crash_log), 'w', encoding='utf-8')
        sys.stdout = sys.stderr
    except Exception:
        pass


logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


def _init_console_logging() -> None:
    if any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        return
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(handler)


def process_pallet_file(path: Path) -> Optional[pd.DataFrame]:
    """Прочитать файл pallet, вставить столбец 'Код палета', удалить строки,
    где значение кода палета встречается в других столбцах (кроме нового столбца).

    Возвращает обработанный DataFrame или None, если файл пуст или невалиден.
    """
    try:
        # Проверка наличия xlrd для .xls файлов
        if path.suffix.lower() == ".xls":
            try:
                import xlrd
            except ImportError:
                logger.error(f"Для чтения .xls файлов требуется пакет xlrd.\n  → Рекомендация: Выполните в терминале: pip install xlrd")
                return None
            engine = "xlrd"
        else:
            engine = "openpyxl"
        logger.info(f"Читаю {path}")
        df = pd.read_excel(path, engine=engine, dtype=str)
    except Exception as e:
        logger.error(f"Не удалось прочитать {path}: {e}\n  → Рекомендация: Убедитесь, что файл существует, не повреждён и имеет формат .xlsx или .xls с установленным xlrd")
        return None

    if df.empty:
        logger.warning(f"Файл {path} пуст — пропускаю.\n  → Рекомендация: Проверьте, что файл содержит данные на первом листе.")
        return None

    # Берём последнюю строку и значение из первого столбца
    try:
        last_row = df.iloc[-1]
    except Exception as e:
        logger.error(f"Не удалось получить последнюю строку в {path}: {e}\n  → Рекомендация: Проверьте структуру файла — он должен содержать хотя бы одну строку данных.")
        return None

    try:
        kod_paleta = str(last_row.iloc[0])
    except Exception as e:
        logger.error(f"Не удалось получить значение кода палета в {path}: {e}\n  → Рекомендация: Убедитесь, что последняя строка файла содержит код палета в первом столбце.")
        return None

    # Добавляем новый столбец после 'Код упаковки'
    cols = list(df.columns)
    insert_after = None
    try:
        insert_after = cols.index("Код упаковки")
    except ValueError:
        # Если столбца нет — попробуем вставить после 2-го столбца, иначе в конец
        logger.warning(f"В {path} нет столбца 'Код упаковки' — вставляю 'Код палета' в позицию 2 или в конец.")
        insert_after = min(1, len(cols) - 1)

    new_col_name = "Код палета"
    # Если столбец уже есть, перезапишем его
    if new_col_name in cols:
        df[new_col_name] = kod_paleta
    else:
        # Вставляем столбец в нужную позицию
        left = cols[: insert_after + 1]
        right = cols[insert_after + 1 :]
        df = df.reindex(columns=left + [new_col_name] + right)
        df[new_col_name] = kod_paleta

    # Удаляем строки, где код палета встречается в других столбцах (кроме 'Код палета')
    def row_contains_pal(row: pd.Series) -> bool:
        # Исключаем столбец 'Код палета' и проверяем появления подстроки
        for c in row.index:
            if c == new_col_name:
                continue
            try:
                if pd.isna(row[c]):
                    continue
                if kod_paleta in str(row[c]):
                    return True
            except Exception:
                continue
        return False

    mask = df.apply(row_contains_pal, axis=1)
    # Удаляем такие строки (включая последнюю)
    cleaned_df = df[~mask].copy()
    # Сброс индекса для аккуратности
    cleaned_df.reset_index(drop=True, inplace=True)

    logger.info(f"В {path.name}: извлечён код палета '{kod_paleta}', удалено {mask.sum()} строк.")
    return cleaned_df


def merge_dataframes(dfs: List[pd.DataFrame]) -> pd.DataFrame:
    """Объединяет список DataFrame-ов последовательно, сохраняя заголовки один раз.
    Приводит к требуемому порядку столбцов.
    """
    if not dfs:
        return pd.DataFrame()

    desired_cols = [
        "Код маркировки",
        "Код упаковки",
        "Код палета",
        "Номенклатура",
        "Номер короба",
    ]

    normalized = []
    for df in dfs:
        # Reindex, добавляя отсутствующие колонки как NaN
        normalized.append(df.reindex(columns=desired_cols))

    merged = pd.concat(normalized, ignore_index=True, sort=False)
    return merged


def clean_parentheses(df: pd.DataFrame) -> pd.DataFrame:
    """Удаляет символы '(' и ')' из первых трёх столбцов, если они существуют."""
    cols = ["Код маркировки", "Код упаковки", "Код палета"]
    for c in cols:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.replace("(", "", regex=False).str.replace(")", "", regex=False)
    return df


def load_specification(path: Path) -> Optional[pd.DataFrame]:
    """Загружает спецификацию — вторую страницу (sheet index 1).

    Ожидаемые столбцы: Pallet number, Order ID, Product name, MFD, BBD, QTY PCS,
    QTY a BOX, QTY BOXES, Volume Lit
    """
    try:
        # Проверка наличия xlrd для .xls файлов
        if path.suffix.lower() == ".xls":
            try:
                import xlrd
            except ImportError:
                logger.error(f"Для чтения .xls файлов требуется пакет xlrd.\n  → Рекомендация: Выполните в терминале: pip install xlrd")
                return None
            engine = "xlrd"
        else:
            engine = "openpyxl"
        logger.info(f"Читаю спецификацию {path} (вторая страница)")
        spec = pd.read_excel(path, engine=engine, sheet_name=1)
        if spec.empty:
            logger.warning("Спецификация пуста.")
            return None
        # Приведём имя столбца 'Pallet number' к строкам
        if "Pallet number" in spec.columns:
            spec["Pallet number"] = spec["Pallet number"].astype(str).str.strip()
        return spec
    except Exception as e:
        logger.error(f"Не удалось загрузить спецификацию: {e}\n  → Рекомендация: Убедитесь, что файл спецификации содержит вторую страницу (лист) с данными.")
        return None


def enrich_with_spec(df: pd.DataFrame, spec: pd.DataFrame) -> pd.DataFrame:
    """Добавляет столбцы MFD и BBD на основе номера палета, извлечённого из 'Номер короба'."""
    if df.empty:
        return df

    # Создаём карту: Pallet number -> {MFD:..., BBD:...}
    mapping: Dict[str, Dict[str, str]] = {}
    if spec is not None and "Pallet number" in spec.columns:
        for _, row in spec.iterrows():
            key = str(row.get("Pallet number", "")).strip()
            mapping.setdefault(key, {"MFD": None, "BBD": None})
            mapping[key]["MFD"] = row.get("MFD")
            mapping[key]["BBD"] = row.get("BBD")

    # Вытащим номер палета из 'Номер короба' (до дефиса)
    def extract_pallet_number(value: str) -> str:
        try:
            s = str(value)
            return s.split("-")[0].strip()
        except Exception:
            return ""

    pallet_numbers = df.get("Номер короба", pd.Series([""] * len(df))).astype(str).apply(extract_pallet_number)

    mfd_list = []
    bbd_list = []
    for pn in pallet_numbers:
        info = mapping.get(pn)
        if info:
            mfd_list.append(info.get("MFD"))
            bbd_list.append(info.get("BBD"))
        else:
            mfd_list.append(None)
            bbd_list.append(None)

    df["MFD"] = mfd_list
    df["BBD"] = bbd_list
    return df


def format_date_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Приводит столбцы дат к формату ДД.ММ.ГГГГ без времени."""
    for c in cols:
        if c not in df.columns:
            continue
        # Преобразуем в datetime, затем в нужный формат
        dt = pd.to_datetime(df[c], errors="coerce")
        df[c] = dt.dt.strftime("%d.%m.%Y")
    return df


def validate_specification(merged_df: pd.DataFrame, spec_df: pd.DataFrame, pallet_files_count: int) -> None:
    """Строгая валидация спецификации.

    Raises Exception при любой несоответствующей проверке.
    """
    logger.info("Начинаю валидацию спецификации...")

    # Проверки на наличие колонок
    required_merged_cols = [
        "Код маркировки",
        "Код упаковки",
        "Код палета",
        "Номенклатура",
        "Номер короба",
    ]
    for c in required_merged_cols:
        if c not in merged_df.columns:
            raise Exception(f"Ошибка: в объединённом файле отсутствует столбец '{c}'")

    required_spec_cols = [
        "Pallet number",
        "Order ID",
        "Product name",
        "MFD",
        "BBD",
        "QTY PCS",
        "QTY a BOX",
        "QTY BOXES",
        "Volume, Lit",
    ]
    for c in required_spec_cols:
        if c not in spec_df.columns:
            raise Exception(f"Ошибка: в файле спецификации отсутствует столбец '{c}'")

    # Шаг 1: проверка количества палетов (строк в спецификации)
    spec_count = len(spec_df)
    if spec_count != pallet_files_count:
        raise Exception("Ошибка: количество палетов не совпадает со спецификацией")

    # Подготовка: приводим 'Pallet number' к строковому виду без пробелов
    spec_df_local = spec_df.copy()
    spec_df_local["Pallet number"] = spec_df_local["Pallet number"].astype(str).str.strip()

    # Шаг 2 и 3: по каждому палету в объединённом DataFrame
    # Извлекаем номер палета из 'Номер короба' (до дефиса)
    def extract_pn(val: str) -> str:
        try:
            return str(val).split("-")[0].strip()
        except Exception:
            return ""

    merged_pn = merged_df.get("Номер короба", pd.Series([""] * len(merged_df))).astype(str).apply(extract_pn)
    merged_df_local = merged_df.copy()
    merged_df_local = merged_df_local.assign(_pallet_number_extracted=merged_pn)

    unique_pallets = merged_df_local["_pallet_number_extracted"].unique()

    for pn in unique_pallets:
        if pn == "":
            raise Exception("Ошибка: найден пустой или некорректный номер палета в 'Номер короба'")

        # Строки для данного палета
        rows = merged_df_local[merged_df_local["_pallet_number_extracted"] == pn]
        if rows.empty:
            raise Exception(f"Ошибка: палет {pn} отсутствует в объединённом файле")

        # Уникальные номенклатуры
        unique_products = rows["Номенклатура"].dropna().astype(str).str.strip().unique()
        if len(unique_products) == 0:
            raise Exception(f"Ошибка: пустая номенклатура для палета {pn}")
        if len(unique_products) > 1:
            raise Exception(f"Ошибка: для палета {pn} найдено более одной номенклатуры")

        merged_product = unique_products[0]

        # Найдём палет в спецификации
        spec_rows = spec_df_local[spec_df_local["Pallet number"] == pn]
        if spec_rows.empty:
            raise Exception(f"Ошибка: палет {pn} отсутствует в спецификации")

        # Получаем значение Product name из спецификации (проверяем однозначность)
        spec_products = spec_rows["Product name"].dropna().astype(str).str.strip().unique()
        if len(spec_products) == 0:
            raise Exception(f"Ошибка: отсутствует Product name в спецификации для палета {pn}")
        if len(spec_products) > 1:
            raise Exception(f"Ошибка: в спецификации для палета {pn} указано более одного Product name")

        spec_product = spec_products[0]

        # Сравнение (строгое)
        if spec_product != merged_product:
            raise Exception(f"Ошибка: несоответствие продукта для палета {pn} — неверная спецификация")

    logger.info("Спецификация успешно проверена")


def get_order_id(spec_df: pd.DataFrame) -> str:
    """Извлекает уникальный Order ID из спецификации.

    Raises Exception если Order ID отсутствует или неоднозначен.
    """
    if "Order ID" not in spec_df.columns:
        raise Exception("В спецификации отсутствует столбец 'Order ID'.\n  → Рекомендация: Убедитесь, что на втором листе спецификации есть столбец 'Order ID'.")

    unique_ids = spec_df["Order ID"].dropna().astype(str).str.strip().unique()
    unique_ids = [x for x in unique_ids if x != "" and x.lower() != "nan"]

    if len(unique_ids) == 0:
        raise Exception("В спецификации отсутствует Order ID.\n  → Рекомендация: Убедитесь, что во всех строках есть Order ID.")
    if len(unique_ids) > 1:
        raise Exception(f"В спецификации найдено несколько Order ID: {', '.join(unique_ids)}.\n  → Рекомендация: Order ID должен быть одинаковым для всех строк спецификации.")

    order_id = unique_ids[0]
    logger.info(f"Order ID: {order_id}")
    return order_id


def generate_output_filename(order_id: str, output_dir: Path) -> Path:
    """Формирует путь к итоговому файлу: <output_dir>/ОТЧЕТ.<OrderID>.xls"""
    safe_name = order_id.strip().replace("/", "_").replace("\\", "_")
    filename = f"ОТЧЕТ. {safe_name}.xls"
    return output_dir / filename


def save_output(df: pd.DataFrame, out_path: Path) -> bool:
    try:
        # Принудительно сохраняем в .xls
        if out_path.suffix.lower() != ".xls":
            out_path = out_path.with_suffix(".xls")
        # Проверка наличия xlwt
        try:
            import xlwt
        except Exception:
            logger.error("Для сохранения .xls требуется пакет xlwt.\n  → Рекомендация: Выполните в терминале: pip install xlwt")
            return False
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Sheet1")

        # Стиль для заголовков: жирный шрифт + выравнивание по центру
        header_style = xlwt.XFStyle()
        header_font = xlwt.Font()
        header_font.bold = True
        header_style.font = header_font
        header_align = xlwt.Alignment()
        header_align.horz = xlwt.Alignment.HORZ_CENTER
        header_style.alignment = header_align

        # Заголовки с применением стиля
        for col_idx, col_name in enumerate(df.columns):
            sheet.write(0, col_idx, col_name, header_style)

        # Подготовим ширины столбцов: вычислим максимальную длину строки в столбце
        # (с учётом заголовка) и установим ширину через sheet.col(...).width.
        # Ограничим ширину до 255 символов (ограничение Excel).
        max_chars_per_col = []
        for col_idx, col_name in enumerate(df.columns):
            max_len = len(str(col_name))
            for val in df.iloc[:, col_idx].fillna(""):
                try:
                    l = len(str(val))
                except Exception:
                    l = 0
                if l > max_len:
                    max_len = l
            # Небольшой запас на визуализацию
            max_len = min(max_len + 2, 255)
            max_chars_per_col.append(max_len)

        for col_idx, max_chars in enumerate(max_chars_per_col):
            try:
                sheet.col(col_idx).width = 256 * int(max_chars)
            except Exception:
                # Если чего-то пойдёт не так — просто пропустим установку ширины
                pass

        # Данные
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):
            for col_idx, value in enumerate(row):
                sheet.write(row_idx, col_idx, "" if pd.isna(value) else value)

        workbook.save(str(out_path))
        logger.info(f"Результат сохранён в {out_path}")
        return True
    except Exception as e:
        logger.error(f"Не удалось сохранить файл {out_path}: {e}\n  → Рекомендация: Проверьте, что путь к файлу корректен и нет открытого файла с таким именем.")
        # Удалить частично созданный файл, если он существует
        if out_path.exists():
            try:
                out_path.unlink()
            except Exception:
                pass  # Игнорируем ошибки удаления
        return False

def run_pipeline(base_dir: Path, spec_path: Optional[Path], out_path: Path) -> bool:
    logger.info(f"Базовая директория: {base_dir}")

    if spec_path and spec_path == out_path:
        logger.error("Файл спецификации не может быть тем же, что и выходной файл.")
        return False

    processed_dfs: List[pd.DataFrame] = []

    try:
        # Шаг 1: обработать отдельные Pallet i.xlsx
        for i in range(1, 31):
            fname = f"Pallet {i}.xlsx"
            p = base_dir / fname
            if not p.exists():
                logger.debug(f"Файл {fname} не найден — пропускаю.")
                continue
            df = process_pallet_file(p)
            if df is not None and not df.empty:
                processed_dfs.append(df)

        # Шаг 2: объединение
        if not processed_dfs:
            logger.error("Нет обработанных файлов для объединения.\n  → Рекомендация: Убедитесь, что в выбранной папке есть файлы 'Pallet 1.xlsx' .. 'Pallet 30.xlsx'.")
            return False

        merged = merge_dataframes(processed_dfs)
        logger.info(f"Объединено {len(processed_dfs)} файлов, итоговых строк: {len(merged)}")

        # Шаг 3: загрузка спецификации (для валидации)
        spec_df = None
        if spec_path:
            if not spec_path.exists():
                logger.error("Файл спецификации не найден. Обработка остановлена.")
                return False
            spec_df = load_specification(spec_path)
            if spec_df is None:
                logger.error("Не удалось загрузить спецификацию. Обработка остановлена.")
                return False

            # Выполняем строгую валидацию спецификации на исходных (неизменённых) данных
            try:
                validate_specification(merged, spec_df, pallet_files_count=len(processed_dfs))
            except Exception as e:
                logger.error(str(e))
                return False

        # Автоматическое формирование имени файла из Order ID
        if spec_df is not None:
            try:
                order_id = get_order_id(spec_df)
                out_path = generate_output_filename(order_id, out_path.parent)
                logger.info(f"Имя выходного файла: {out_path.name}")
            except Exception as e:
                logger.error(str(e))
                return False

        # Шаг 4: очистка первых трёх столбцов
        merged = clean_parentheses(merged)

        # Шаг 5: обогащение (если спецификация была загружена)
        if spec_df is not None:
            merged = enrich_with_spec(merged, spec_df)

        # Шаг 5.1: форматирование дат
        merged = format_date_columns(merged, ["MFD", "BBD"])

        # Шаг 6: сохранение (только если ошибок не было)
        save_ok = save_output(merged, out_path)
        if not save_ok:
            logger.error("Сохранение результата завершилось с ошибкой. Файл не создан.\n  → Рекомендация: Закройте открытый выходной файл и попробуйте снова.")
            return False

        return True
    except Exception as e:
        logger.exception(f"Фатальная ошибка при выполнении: {e}\n  → Рекомендация: Проверьте входные файлы и повторите попытку. Если проблема повторяется — обратитесь к разработчику.")
        return False


class TextHandler(logging.Handler):
    """Логгер для вывода в Text-виджет Tkinter."""

    def __init__(self, text_widget: tk.Text) -> None:
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", msg + "\n")
        self.text_widget.configure(state="disabled")
        self.text_widget.see("end")


# Глобальный коллектор ошибок — заполняется из логгера
error_collector: List[str] = []
warning_collector: List[str] = []


class ErrorCollectorHandler(logging.Handler):
    """Собирает ERROR-сообщения в глобальный список для показа пользователю."""

    def __init__(self) -> None:
        super().__init__(level=logging.ERROR)

    def emit(self, record: logging.LogRecord) -> None:
        error_collector.append(self.format(record))


class WarningCollectorHandler(logging.Handler):
    """Собирает WARNING-сообщения в глобальный список."""

    def __init__(self) -> None:
        super().__init__(level=logging.WARNING)

    def emit(self, record: logging.LogRecord) -> None:
        if record.levelno == logging.WARNING:
            warning_collector.append(self.format(record))


def _show_error_window(parent: tk.Tk, title: str, details: str) -> None:
    """Показывает красивое окно ошибки с подробностями, разделяя ошибки и рекомендации."""
    win = tk.Toplevel(parent)
    win.title("⚠️ Ошибка")
    win.geometry("620x400")
    win.resizable(True, True)
    win.grab_set()

    frame = ttk.Frame(win, padding=16)
    frame.pack(fill="both", expand=True)

    # Заголовок ошибки
    lbl = ttk.Label(frame, text=title, font=("TkDefaultFont", 14, "bold"),
                    foreground="red", wraplength=580)
    lbl.pack(anchor="w", pady=(0, 8))

    # Разделяем детали на ошибки и рекомендации
    errors = []
    recommendations = []
    for line in details.split('\n'):
        if '→ Рекомендация:' in line:
            rec = line.split('→ Рекомендация:', 1)[1].strip()
            recommendations.append(rec)
        else:
            errors.append(line)

    # Ошибки
    if errors:
        err_frame = ttk.Frame(frame)
        err_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(err_frame, text="Ошибки:", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        err_txt = tk.Text(err_frame, wrap="word", height=6, font=("TkFixedFont", 10))
        err_txt.insert("1.0", '\n'.join(errors))
        err_txt.configure(state="disabled")
        err_scroll = ttk.Scrollbar(err_frame, command=err_txt.yview)
        err_txt.configure(yscrollcommand=err_scroll.set)
        err_txt.pack(side="left", fill="both", expand=True)
        err_scroll.pack(side="right", fill="y")

    # Рекомендации
    if recommendations:
        rec_frame = ttk.Frame(frame)
        rec_frame.pack(fill="x")
        ttk.Label(rec_frame, text="Рекомендации:", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        rec_txt = tk.Text(rec_frame, wrap="word", height=6, font=("TkFixedFont", 10))
        rec_txt.insert("1.0", '\n'.join(recommendations))
        rec_txt.configure(state="disabled")
        rec_scroll = ttk.Scrollbar(rec_frame, command=rec_txt.yview)
        rec_txt.configure(yscrollcommand=rec_scroll.set)
        rec_txt.pack(side="left", fill="both", expand=True)
        rec_scroll.pack(side="right", fill="y")

    ttk.Button(win, text="Закрыть", command=win.destroy).pack(pady=8)


def launch_gui() -> None:
    # Проверка зависимостей
    try:
        import xlrd
    except ImportError:
        messagebox.showerror("Отсутствует зависимость", "Требуется пакет xlrd для чтения Excel файлов.\nУстановите: pip install xlrd")
        return
    try:
        import openpyxl
    except ImportError:
        messagebox.showerror("Отсутствует зависимость", "Требуется пакет openpyxl для чтения .xlsx файлов.\nУстановите: pip install openpyxl")
        return
    try:
        import xlwt
    except ImportError:
        messagebox.showerror("Отсутствует зависимость", "Требуется пакет xlwt для сохранения .xls файлов.\nУстановите: pip install xlwt")
        return

    root = tk.Tk()
    root.title("Pallet Merger")
    root.geometry("780x320")
    root.minsize(720, 280)

    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")

    main_frame = ttk.Frame(root, padding=16)
    main_frame.pack(fill="both", expand=True)

    # Пути
    base_dir_var = tk.StringVar(value=str(Path.cwd()))
    spec_path_var = tk.StringVar(value="")
    out_path_var = tk.StringVar(value=str(Path.cwd() / "Merged_Pallets.xls"))

    def choose_base_dir() -> None:
        path = filedialog.askdirectory(title="Выберите папку с Pallet-файлами")
        if path:
            base_dir_var.set(path)
            out_path_var.set(str(Path(path) / "Merged_Pallets.xls"))

    def choose_spec_file() -> None:
        path = filedialog.askopenfilename(
            title="Выберите файл спецификации",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            spec_path_var.set(path)

    def choose_output_file() -> None:
        path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".xls",
            filetypes=[("Excel files", "*.xls")],
        )
        if path:
            out_path_var.set(path)

    def open_output_file() -> None:
        p = Path(out_path_var.get())
        if not p.exists():
            messagebox.showerror("Ошибка", "Выходной файл не найден.")
            return
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(p)])
            elif sys.platform.startswith("win"):
                os.startfile(str(p))
            else:
                subprocess.run(["xdg-open", str(p)])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")

    # ── Панель ошибок (скрыта по умолчанию) ──
    error_panel = tk.Frame(main_frame, bg="#FFDDDD", bd=2, relief="groove")
    # НЕ делаем grid — скрыта
    error_title_lbl = tk.Label(error_panel, text="⚠ Ошибка обработки",
                               font=("TkDefaultFont", 13, "bold"),
                               fg="#CC0000", bg="#FFDDDD", anchor="w")
    error_title_lbl.pack(fill="x", padx=10, pady=(8, 2))
    error_msg_lbl = tk.Label(error_panel, text="", fg="#880000", bg="#FFDDDD",
                             font=("TkDefaultFont", 11), anchor="w",
                             justify="left", wraplength=700)
    error_msg_lbl.pack(fill="x", padx=10, pady=(0, 2))
    error_rec_lbl = tk.Label(error_panel, text="", fg="#555555", bg="#FFDDDD",
                             font=("TkDefaultFont", 10, "italic"), anchor="w",
                             justify="left", wraplength=700)
    error_rec_lbl.pack(fill="x", padx=10, pady=(0, 8))

    # ── Панель успеха (скрыта по умолчанию) ──
    success_panel = tk.Frame(main_frame, bg="#DDFFDD", bd=2, relief="groove")
    success_lbl = tk.Label(success_panel, text="✔ Обработка завершена успешно!",
                           font=("TkDefaultFont", 12, "bold"),
                           fg="#007700", bg="#DDFFDD")
    success_lbl.pack(fill="x", padx=10, pady=8)

    def show_error_panel(message: str, recommendation: str = "") -> None:
        success_panel.grid_forget()
        error_msg_lbl.config(text=message)
        error_rec_lbl.config(text=f"Рекомендация: {recommendation}" if recommendation else "")
        error_panel.grid(row=100, column=0, columnspan=3, sticky="we", pady=(8, 0))

    def hide_error_panel() -> None:
        error_panel.grid_forget()

    def show_status(message: str) -> None:
        hide_error_panel()
        success_lbl.config(text=message)
        success_panel.grid(row=100, column=0, columnspan=3, sticky="we", pady=(8, 0))

    def hide_status() -> None:
        success_panel.grid_forget()

    def on_run() -> None:
        base_dir = Path(base_dir_var.get())
        spec_path = Path(spec_path_var.get()) if spec_path_var.get().strip() else None
        out_path = Path(out_path_var.get())

        hide_error_panel()
        hide_status()

        if not base_dir.exists():
            show_error_panel("Папка с Pallet-файлами не найдена.",
                             "Убедитесь, что указанная папка существует и содержит файлы Pallet.")
            return

        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")

        # Собираем ошибки и предупреждения из логов
        error_collector.clear()
        warning_collector.clear()
        success = run_pipeline(base_dir=base_dir, spec_path=spec_path, out_path=out_path)

        if success and not error_collector:
            show_status("✔ Обработка завершена успешно!")
        else:
            # Разделяем ошибки и рекомендации
            raw = "\n".join(error_collector) if error_collector else "Неизвестная ошибка."
            errors = []
            recs = []
            for line in raw.split("\n"):
                if "→ Рекомендация:" in line:
                    recs.append(line.split("→ Рекомендация:", 1)[1].strip())
                elif line.strip():
                    errors.append(line.strip())
            show_error_panel("\n".join(errors), "\n".join(recs))
            # Автоматически показать логи при ошибке
            if not log_visible.get():
                toggle_logs()

    # Верхние поля
    row = 0
    ttk.Label(main_frame, text="Папка с Pallet-файлами").grid(row=row, column=0, sticky="w")
    ttk.Entry(main_frame, textvariable=base_dir_var, width=70).grid(row=row, column=1, sticky="we", padx=8)
    ttk.Button(main_frame, text="Выбрать", command=choose_base_dir).grid(row=row, column=2)

    row += 1
    ttk.Label(main_frame, text="Файл спецификации").grid(row=row, column=0, sticky="w", pady=(8, 0))
    ttk.Entry(main_frame, textvariable=spec_path_var, width=70).grid(row=row, column=1, sticky="we", padx=8, pady=(8, 0))
    ttk.Button(main_frame, text="Выбрать", command=choose_spec_file).grid(row=row, column=2, pady=(8, 0))

    row += 1
    ttk.Label(main_frame, text="Выходной файл").grid(row=row, column=0, sticky="w", pady=(8, 0))
    ttk.Entry(main_frame, textvariable=out_path_var, width=70).grid(row=row, column=1, sticky="we", padx=8, pady=(8, 0))
    ttk.Button(main_frame, text="Выбрать", command=choose_output_file).grid(row=row, column=2, pady=(8, 0))

    row += 1
    ttk.Button(main_frame, text="Запустить", command=on_run).grid(row=row, column=0, columnspan=3, sticky="we", pady=12)

    # Кнопки: Открыть файл + Показать логи
    btn_frame = ttk.Frame(main_frame)
    btn_frame.grid(row=row + 1, column=0, columnspan=3, sticky="we")
    ttk.Button(btn_frame, text="Открыть файл", command=open_output_file).pack(side="right", padx=4)

    log_visible = tk.BooleanVar(value=False)
    toggle_btn_text = tk.StringVar(value="Показать логи ▼")

    # Лог (скрыт по умолчанию)
    log_frame = ttk.Frame(main_frame)
    # НЕ делаем grid — скрыто
    log_frame.rowconfigure(0, weight=1)
    log_frame.columnconfigure(0, weight=1)

    log_text = tk.Text(log_frame, height=12, wrap="word", state="disabled")
    log_scroll = ttk.Scrollbar(log_frame, command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    log_text.grid(row=0, column=0, sticky="nsew")
    log_scroll.grid(row=0, column=1, sticky="ns")

    def toggle_logs() -> None:
        if log_visible.get():
            log_frame.grid_forget()
            log_visible.set(False)
            toggle_btn_text.set("Показать логи ▼")
            root.geometry("780x320")
        else:
            log_frame.grid(row=row + 3, column=0, columnspan=3, sticky="nsew")
            log_visible.set(True)
            toggle_btn_text.set("Скрыть логи ▲")
            root.geometry("780x560")

    ttk.Button(btn_frame, textvariable=toggle_btn_text, command=toggle_logs).pack(side="right", padx=4)

    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(row + 3, weight=1)

    # Логи в GUI
    gui_handler = TextHandler(log_text)
    gui_handler.setLevel(logging.INFO)
    logger.addHandler(gui_handler)

    # Коллектор ошибок для статуса
    err_handler = ErrorCollectorHandler()
    err_handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(err_handler)

    # Коллектор предупреждений
    warn_handler = WarningCollectorHandler()
    warn_handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(warn_handler)

    root.mainloop()


def main() -> None:
    # Проверка зависимостей
    try:
        import xlrd
    except ImportError:
        print("Ошибка: Требуется пакет xlrd для чтения Excel файлов. Установите: pip install xlrd", file=sys.stderr)
        sys.exit(1)
    try:
        import openpyxl
    except ImportError:
        print("Ошибка: Требуется пакет openpyxl для чтения .xlsx файлов. Установите: pip install openpyxl", file=sys.stderr)
        sys.exit(1)
    try:
        import xlwt
    except ImportError:
        print("Ошибка: Требуется пакет xlwt для сохранения .xls файлов. Установите: pip install xlwt", file=sys.stderr)
        sys.exit(1)

    _init_console_logging()
    parser = argparse.ArgumentParser(description="Pallet merger")
    parser.add_argument("--cli", action="store_true", help="Запуск без GUI")
    parser.add_argument("--base-dir", type=str, help="Папка с Pallet-файлами")
    parser.add_argument("--spec", type=str, help="Путь к файлу спецификации")
    parser.add_argument("--out", type=str, help="Путь к выходному файлу")
    args = parser.parse_args()

    if args.cli:
        base_dir = Path(args.base_dir) if args.base_dir else Path.cwd()
        spec_path = Path(args.spec) if args.spec else None
        out_path = Path(args.out) if args.out else base_dir / "Merged_Pallets.xls"
        if spec_path and spec_path == out_path:
            print("Ошибка: Файл спецификации не может быть тем же, что и выходной файл.", file=sys.stderr)
            sys.exit(1)
        ok = run_pipeline(base_dir=base_dir, spec_path=spec_path, out_path=out_path)
        if not ok:
            sys.exit(1)
    else:
        launch_gui()


if __name__ == "__main__":
    main()
