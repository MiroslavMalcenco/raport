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

import pandas as pd


logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


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
        spec = pd.read_excel(path, engine="openpyxl", sheet_name=1, dtype=str)
        if spec.empty:
            logger.warning("Спецификация пуста.")
            return None
        # Приведём имя столбца 'Pallet number' к строкам
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


def save_output(df: pd.DataFrame, out_path: Path) -> None:
    try:
        df.to_excel(out_path, index=False, engine="openpyxl")
        logger.info(f"Результат сохранён в {out_path}")
    except Exception as e:
        logger.error(f"Не удалось сохранить файл {out_path}: {e}")


def main() -> None:
    base_dir = Path.cwd()
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
        sys.exit(1)

    merged = merge_dataframes(processed_dfs)
    logger.info(f"Объединено {len(processed_dfs)} файлов, итоговых строк: {len(merged)}")

    # Шаг 3: очистка первых трёх столбцов
    merged = clean_parentheses(merged)

    # Шаг 4: загрузка спецификации от пользователя
    spec_path_input = input("Введите путь к файлу спецификации (Excel): ").strip()
    spec_path = Path(spec_path_input)
    if not spec_path.exists():
        logger.error(f"Файл спецификации {spec_path} не найден. Обогащение пропущено.")
        spec_df = None
    else:
        spec_df = load_specification(spec_path)

    # Шаг 5: обогащение
    merged = enrich_with_spec(merged, spec_df)

    # Шаг 6: сохранение
    out = base_dir / "Merged_Pallets.xlsx"
    save_output(merged, out)


if __name__ == "__main__":
    main()
