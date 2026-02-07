#!/usr/bin/env python3
"""merge_pallets.py

Скрипт для обработки файлов Pallet 1.xlsx .. Pallet 30.xlsx, объединения,
очистки данных и обогащения по спецификации.

Требования: pandas, openpyxl
"""
from pathlib import Path
import logging
import sys
from typing import List, Dict, Optional

import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


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
        logger.info(f"Читаю {path}")
        df = pd.read_excel(path, engine="openpyxl", dtype=str)
    except Exception as e:
        logger.error(f"Не удалось прочитать {path}: {e}")
        return None

    if df.empty:
        logger.warning(f"Файл {path} пуст — пропускаю.")
        return None

    # Берём последнюю строку и значение из первого столбца
    try:
        last_row = df.iloc[-1]
    except Exception as e:
        logger.error(f"Не удалось получить последнюю строку в {path}: {e}")
        return None

    try:
        kod_paleta = str(last_row.iloc[0])
    except Exception as e:
        logger.error(f"Не удалось получить значение кода палета в {path}: {e}")
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
        logger.info(f"Читаю спецификацию {path} (вторая страница)")
        spec = pd.read_excel(path, engine="openpyxl", sheet_name=1)
        if spec.empty:
            logger.warning("Спецификация пуста.")
            return None
        # Приведём имя столбца 'Pallet number' к строкам
        if "Pallet number" in spec.columns:
            spec["Pallet number"] = spec["Pallet number"].astype(str).str.strip()
        return spec
    except Exception as e:
        logger.error(f"Не удалось загрузить спецификацию: {e}")
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


def save_output(df: pd.DataFrame, out_path: Path) -> None:
    try:
        # Принудительно сохраняем в .xls
        if out_path.suffix.lower() != ".xls":
            out_path = out_path.with_suffix(".xls")
        # Проверка наличия xlwt
        try:
            import xlwt
        except Exception:
            logger.error("Для сохранения .xls требуется пакет xlwt. Установите: pip install xlwt")
            return
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Sheet1")

        # Заголовки
        for col_idx, col_name in enumerate(df.columns):
            sheet.write(0, col_idx, col_name)

        # Данные
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):
            for col_idx, value in enumerate(row):
                sheet.write(row_idx, col_idx, "" if pd.isna(value) else value)

        workbook.save(str(out_path))
        logger.info(f"Результат сохранён в {out_path}")
    except Exception as e:
        logger.error(f"Не удалось сохранить файл {out_path}: {e}")

def run_pipeline(base_dir: Path, spec_path: Optional[Path], out_path: Path) -> None:
    logger.info(f"Базовая директория: {base_dir}")

    processed_dfs: List[pd.DataFrame] = []

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
        logger.error("Нет обработанных файлов для объединения. Завершаю работу.")
        return

    merged = merge_dataframes(processed_dfs)
    logger.info(f"Объединено {len(processed_dfs)} файлов, итоговых строк: {len(merged)}")

    # Шаг 3: очистка первых трёх столбцов
    merged = clean_parentheses(merged)

    # Шаг 4: загрузка спецификации
    spec_df = None
    if spec_path and spec_path.exists():
        spec_df = load_specification(spec_path)
    else:
        logger.warning("Файл спецификации не указан или не найден. Обогащение пропущено.")

    # Шаг 5: обогащение
    merged = enrich_with_spec(merged, spec_df)

    # Шаг 5.1: форматирование дат
    merged = format_date_columns(merged, ["MFD", "BBD"])

    # Шаг 6: сохранение
    save_output(merged, out_path)


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


def launch_gui() -> None:
    root = tk.Tk()
    root.title("Pallet Merger")
    root.geometry("780x520")
    root.minsize(720, 480)

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

    def on_run() -> None:
        base_dir = Path(base_dir_var.get())
        spec_path = Path(spec_path_var.get()) if spec_path_var.get().strip() else None
        out_path = Path(out_path_var.get())

        if not base_dir.exists():
            messagebox.showerror("Ошибка", "Папка с Pallet-файлами не найдена.")
            return

        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")

        run_pipeline(base_dir=base_dir, spec_path=spec_path, out_path=out_path)
        messagebox.showinfo("Готово", "Обработка завершена.")

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

    # Лог
    ttk.Label(main_frame, text="Лог выполнения").grid(row=row + 1, column=0, sticky="w")
    log_frame = ttk.Frame(main_frame)
    log_frame.grid(row=row + 2, column=0, columnspan=3, sticky="nsew")
    log_frame.rowconfigure(0, weight=1)
    log_frame.columnconfigure(0, weight=1)

    log_text = tk.Text(log_frame, height=12, wrap="word", state="disabled")
    log_scroll = ttk.Scrollbar(log_frame, command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    log_text.grid(row=0, column=0, sticky="nsew")
    log_scroll.grid(row=0, column=1, sticky="ns")

    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(row + 2, weight=1)

    # Логи в GUI
    gui_handler = TextHandler(log_text)
    gui_handler.setLevel(logging.INFO)
    logger.addHandler(gui_handler)

    root.mainloop()


def main() -> None:
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
        run_pipeline(base_dir=base_dir, spec_path=spec_path, out_path=out_path)
    else:
        launch_gui()


if __name__ == "__main__":
    main()
