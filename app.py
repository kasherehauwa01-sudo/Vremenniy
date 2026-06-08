"""Streamlit-приложение для преобразования Excel-файлов контроля сроков годности."""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd
import streamlit as st


# Регулярное выражение для поиска годов вида 20XX.
YEAR_PATTERN = re.compile(r"20\d{2}")

# Структура исходного файла по скрину:
# A = код, B = артикул, C = наименование, D = кол-во, E:P = месяцы.
# Индексы pandas начинаются с 0, поэтому B = 1, D = 3, E = 4, ..., P = 15.
ARTICLE_COLUMN_INDEX = 1
QUANTITY_COLUMN_INDEX = 3
REQUIRED_COLUMN_COUNT = 16

MONTH_COLUMNS = {
    4: "01",  # E = ЯНВАРЬ
    5: "02",  # F = ФЕВРАЛЬ
    6: "03",  # G = МАРТ
    7: "04",  # H = АПРЕЛЬ
    8: "05",  # I = МАЙ
    9: "06",  # J = ИЮНЬ
    10: "07",  # K = ИЮЛЬ
    11: "08",  # L = АВГУСТ
    12: "09",  # M = СЕНТЯБРЬ
    13: "10",  # N = ОКТЯБРЬ
    14: "11",  # O = НОЯБРЬ
    15: "12",  # P = ДЕКАБРЬ
}

RESULT_COLUMNS = ["Артикул", "Количество", "Срок годности до"]


class FileProcessingError(Exception):
    """Ошибка, понятная пользователю при чтении или обработке файла."""


def load_file(uploaded_file: BinaryIO) -> pd.DataFrame:
    """
    Загружает Excel-файл в DataFrame.

    Файл читается без строки заголовков (header=None), потому что структура
    задана фиксированными позициями столбцов: B = артикул, D = кол-во,
    E:P = месяцы. Приложение проверяет все листы книги и выбирает первый
    лист с подходящей структурой и найденными годами. Если годов нет ни на
    одном подходящем листе, выбирается первый непустой лист со структурой A:P.
    """
    if uploaded_file is None:
        raise FileProcessingError("Файл не выбран.")

    file_name = getattr(uploaded_file, "name", "")
    file_extension = Path(file_name).suffix.lower()

    if file_extension not in {".xls", ".xlsx"}:
        raise FileProcessingError(
            "Неподдерживаемый формат файла. Загрузите файл с расширением .xls или .xlsx."
        )

    engine = "xlrd" if file_extension == ".xls" else "openpyxl"

    try:
        # Streamlit может переиспользовать объект файла, поэтому перед чтением
        # возвращаем указатель в начало.
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)

        sheets = pd.read_excel(
            uploaded_file,
            header=None,
            sheet_name=None,
            engine=engine,
        )
    except ImportError as exc:
        raise FileProcessingError(
            f"Не установлена библиотека для чтения {file_extension}-файлов: {engine}. "
            "Установите зависимости: streamlit pandas openpyxl xlrd."
        ) from exc
    except Exception as exc:
        raise FileProcessingError(
            "Не удалось прочитать Excel-файл. Проверьте, что файл не повреждён "
            "и соответствует формату .xls или .xlsx."
        ) from exc

    if not sheets:
        raise FileProcessingError("В Excel-файле не найдено листов для обработки.")

    load_logs = [
        f"Загружена книга Excel. Найдено листов: {len(sheets)}.",
        "Проверяются все листы, потому что данные могут находиться не на первой вкладке.",
    ]
    valid_sheets: list[tuple[str, pd.DataFrame, int]] = []
    non_empty_sheet_count = 0

    for sheet_name, sheet_data in sheets.items():
        if sheet_data.empty:
            load_logs.append(f"Лист '{sheet_name}' пропущен: лист пустой.")
            continue

        non_empty_sheet_count += 1

        if sheet_data.shape[1] < REQUIRED_COLUMN_COUNT:
            load_logs.append(
                f"Лист '{sheet_name}' пропущен: найдено столбцов {sheet_data.shape[1]}, "
                f"а требуется минимум {REQUIRED_COLUMN_COUNT} (A:P)."
            )
            continue

        year_count = _count_years_in_month_columns(sheet_data)
        valid_sheets.append((sheet_name, sheet_data, year_count))
        load_logs.append(
            f"Лист '{sheet_name}' подходит по структуре: "
            f"{sheet_data.shape[0]} строк, {sheet_data.shape[1]} столбцов, "
            f"найдено годов в E:P: {year_count}."
        )

    if non_empty_sheet_count == 0:
        raise FileProcessingError("Загруженный файл не содержит данных ни на одном листе.")

    if not valid_sheets:
        raise FileProcessingError(
            "Ни один лист не соответствует ожидаемой структуре. "
            "В файле должно быть минимум 16 столбцов: "
            "A = Код, B = Артикул, C = Наименование, D = Кол-во, E:P = месяцы."
        )

    selected_sheet_name, selected_data, selected_year_count = _select_sheet(valid_sheets)
    selected_data.attrs["source_sheet_name"] = selected_sheet_name
    selected_data.attrs["source_sheet_year_count"] = selected_year_count
    selected_data.attrs["load_logs"] = load_logs + [
        f"Выбран лист '{selected_sheet_name}' для обработки."
    ]

    return selected_data


def _count_years_in_month_columns(data: pd.DataFrame) -> int:
    """Считает количество годов вида 20XX в месячных столбцах E:P."""
    year_count = 0

    for column_index in MONTH_COLUMNS:
        if column_index >= data.shape[1]:
            continue

        year_count += data.iloc[:, column_index].map(
            lambda value: len(YEAR_PATTERN.findall(_cell_to_text(value)))
        ).sum()

    return int(year_count)


def _select_sheet(
    valid_sheets: list[tuple[str, pd.DataFrame, int]],
) -> tuple[str, pd.DataFrame, int]:
    """Выбирает первый подходящий лист, отдавая приоритет листам с найденными годами."""
    for sheet_name, sheet_data, year_count in valid_sheets:
        if year_count > 0:
            return sheet_name, sheet_data, year_count

    return valid_sheets[0]


def _normalize_quantity(value: object) -> object:
    """
    Возвращает количество из исходного файла или 0, если значение отсутствует.

    Отсутствующим считается значение NaN, пустая строка или строка,
    состоящая только из пробелов.
    """
    if pd.isna(value):
        return 0

    if isinstance(value, str) and value.strip() == "":
        return 0

    return value


def _cell_to_text(value: object) -> str:
    """Преобразует значение ячейки в текст для поиска годов регулярным выражением."""
    if pd.isna(value):
        return ""

    return str(value)


def transform_data(data: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    Преобразует исходную таблицу в таблицу со сроками годности и логи обработки.

    Для каждой строки исходного файла просматриваются месячные столбцы E:P.
    Каждый найденный год вида 20XX превращается в отдельную строку результата
    с датой в строгом формате ММ.ГГГГ.
    """
    if data.shape[1] < REQUIRED_COLUMN_COUNT:
        raise FileProcessingError(
            "Недостаточно столбцов для обработки. Ожидаются столбцы A:P."
        )

    result_rows: list[dict[str, object]] = []
    load_logs = data.attrs.get("load_logs", [])
    source_sheet_name = data.attrs.get("source_sheet_name", "неизвестный лист")
    source_sheet_year_count = data.attrs.get("source_sheet_year_count", 0)
    logs: list[str] = [
        *load_logs,
        "Начата обработка данных.",
        f"Обрабатываемый лист: '{source_sheet_name}'.",
        "Годов найдено на выбранном листе при предварительной проверке: "
        f"{source_sheet_year_count}.",
        f"Исходный размер таблицы: {data.shape[0]} строк, {data.shape[1]} столбцов.",
        "Используемая структура: B = Артикул, D = Количество, E:P = месяцы.",
    ]
    rows_with_years = 0
    rows_without_years = 0
    empty_quantity_replacements = 0
    found_year_count = 0

    # Сортировка не выполняется: порядок обхода повторяет исходный порядок строк,
    # внутри строки — порядок месяцев E:P, внутри ячейки — порядок найденных годов.
    for row_number, row in data.iterrows():
        article = row.iloc[ARTICLE_COLUMN_INDEX]
        raw_quantity = row.iloc[QUANTITY_COLUMN_INDEX]
        quantity = _normalize_quantity(raw_quantity)
        row_has_year = False

        quantity_was_empty = pd.isna(raw_quantity) or (
            isinstance(raw_quantity, str) and raw_quantity.strip() == ""
        )

        for column_index, month in MONTH_COLUMNS.items():
            cell_text = _cell_to_text(row.iloc[column_index])
            years = YEAR_PATTERN.findall(cell_text)

            if years:
                row_has_year = True
                found_year_count += len(years)
                excel_row_number = row_number + 1
                excel_column_letter = _column_index_to_excel_letter(column_index)
                logs.append(
                    "Строка "
                    f"{excel_row_number}, столбец {excel_column_letter}: "
                    f"найдены годы {', '.join(years)}."
                )

            for year in years:
                if quantity_was_empty:
                    empty_quantity_replacements += 1

                result_rows.append(
                    {
                        "Артикул": article,
                        "Количество": quantity,
                        "Срок годности до": f"{month}.{year}",
                    }
                )

        if row_has_year:
            rows_with_years += 1
        else:
            rows_without_years += 1

    logs.extend(
        [
            f"Строк с найденными годами: {rows_with_years}.",
            f"Строк без найденных годов исключено: {rows_without_years}.",
            f"Всего найдено годов: {found_year_count}.",
            "Итоговых строк с количеством, заменённым на 0: "
            f"{empty_quantity_replacements}.",
            f"Итоговых строк сформировано: {len(result_rows)}.",
            "Обработка завершена.",
        ]
    )

    result_data = pd.DataFrame(result_rows, columns=RESULT_COLUMNS)
    return result_data, logs


def _column_index_to_excel_letter(column_index: int) -> str:
    """Преобразует индекс столбца pandas в буквенное обозначение Excel."""
    column_number = column_index + 1
    letters = ""

    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters = chr(65 + remainder) + letters

    return letters


def generate_excel(result_data: pd.DataFrame) -> bytes:
    """
    Формирует XLSX-файл с результатом и возвращает его содержимое в байтах.
    """
    output = BytesIO()

    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_data.to_excel(writer, index=False, sheet_name="Результат")
    except Exception as exc:
        raise FileProcessingError("Не удалось сформировать итоговый XLSX-файл.") from exc

    output.seek(0)
    return output.getvalue()


def generate_zip(processed_files: list[dict[str, object]]) -> bytes:
    """Формирует ZIP-архив с итоговыми XLSX-файлами для массовой загрузки."""
    output = BytesIO()

    try:
        with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            for processed_file in processed_files:
                archive.writestr(
                    str(processed_file["output_file_name"]),
                    processed_file["excel_bytes"],
                )
    except Exception as exc:
        raise FileProcessingError("Не удалось сформировать ZIP-архив с результатами.") from exc

    output.seek(0)
    return output.getvalue()


def _build_output_file_name(input_file_name: str) -> str:
    """Создаёт понятное имя итогового XLSX-файла на основе имени исходного файла."""
    source_stem = Path(input_file_name).stem or "result"
    safe_stem = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_-]+", "_", source_stem).strip("_")

    if not safe_stem:
        safe_stem = "result"

    return f"{safe_stem}_result.xlsx"


def _make_unique_file_name(file_name: str, used_file_names: set[str]) -> str:
    """Возвращает уникальное имя файла, чтобы в ZIP-архиве не было дублей."""
    if file_name not in used_file_names:
        used_file_names.add(file_name)
        return file_name

    path = Path(file_name)
    suffix = path.suffix
    stem = path.stem
    counter = 2

    while True:
        unique_file_name = f"{stem}_{counter}{suffix}"

        if unique_file_name not in used_file_names:
            used_file_names.add(unique_file_name)
            return unique_file_name

        counter += 1


def main() -> None:
    """Запускает интерфейс Streamlit-приложения."""
    st.set_page_config(page_title="Преобразование сроков годности", layout="wide")

    st.title("Преобразование сроков годности")
    st.write(
        "Загрузите один или несколько Excel-файлов контроля сроков годности. "
        "Приложение проверит все вкладки каждой книги и обработает структуру "
        "со столбцами A = Код, B = Артикул, C = Наименование, D = Кол-во, "
        "E:P = месяцы. Для каждого файла будет сформирована таблица с колонками: "
        "Артикул, Количество, Срок годности до."
    )

    uploaded_files = st.file_uploader(
        "Загрузите Excel-файл или несколько файлов (.xls или .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Выберите один или несколько файлов Excel для обработки.")
        return

    st.caption(f"Выбрано файлов: {len(uploaded_files)}.")

    if st.button("Обработать", type="primary"):
        processed_files: list[dict[str, object]] = []
        processing_logs: list[str] = []
        failed_files: list[str] = []
        used_output_file_names: set[str] = set()

        with st.spinner("Идёт обработка файлов..."):
            for file_number, uploaded_file in enumerate(uploaded_files, start=1):
                file_name = getattr(uploaded_file, "name", f"file_{file_number}")
                processing_logs.extend(
                    [
                        "=" * 80,
                        f"Файл {file_number} из {len(uploaded_files)}: {file_name}.",
                    ]
                )

                try:
                    source_data = load_file(uploaded_file)
                    result_data, logs = transform_data(source_data)
                    excel_bytes = generate_excel(result_data)
                    output_file_name = _make_unique_file_name(
                        _build_output_file_name(file_name),
                        used_output_file_names,
                    )

                    processed_files.append(
                        {
                            "input_file_name": file_name,
                            "output_file_name": output_file_name,
                            "result_data": result_data,
                            "excel_bytes": excel_bytes,
                            "logs": logs,
                        }
                    )
                    processing_logs.extend(logs)
                    processing_logs.append(
                        f"Файл '{file_name}' успешно обработан. "
                        f"Итоговых строк: {len(result_data)}."
                    )
                except FileProcessingError as exc:
                    failed_files.append(file_name)
                    processing_logs.append(f"Ошибка в файле '{file_name}': {exc}")
                except Exception as exc:
                    failed_files.append(file_name)
                    processing_logs.append(
                        f"Непредвиденная ошибка в файле '{file_name}': {exc}"
                    )

        zip_bytes = generate_zip(processed_files) if len(processed_files) > 1 else None

        st.session_state["processed_files"] = processed_files
        st.session_state["zip_bytes"] = zip_bytes
        st.session_state["processing_logs"] = processing_logs
        st.session_state["failed_files"] = failed_files

        if processed_files:
            st.success(
                "Обработка завершена. "
                f"Успешно обработано файлов: {len(processed_files)} из {len(uploaded_files)}."
            )

            empty_result_count = sum(
                processed_file["result_data"].empty for processed_file in processed_files
            )
            if empty_result_count:
                st.warning(
                    "Для части файлов в месячных столбцах E:P не найдено годов вида 20XX. "
                    f"Файлов с пустым результатом: {empty_result_count}."
                )
        else:
            st.error("Не удалось обработать ни один файл. Подробности смотрите в логах.")

        if failed_files:
            st.warning(
                "Некоторые файлы не были обработаны: " + ", ".join(failed_files)
            )

    processed_files = st.session_state.get("processed_files", [])
    zip_bytes = st.session_state.get("zip_bytes")
    processing_logs = st.session_state.get("processing_logs")

    if processing_logs:
        st.subheader("Логи обработки")
        st.text_area(
            "Подробный журнал",
            value="\n".join(processing_logs),
            height=320,
            disabled=True,
        )

    if processed_files:
        st.subheader("Предпросмотр результата")

        for processed_file in processed_files:
            result_data = processed_file["result_data"]
            input_file_name = processed_file["input_file_name"]

            with st.expander(
                f"{input_file_name} — первые 100 строк "
                f"(всего строк: {len(result_data)})",
                expanded=len(processed_files) == 1,
            ):
                st.dataframe(result_data.head(100), use_container_width=True)

        if len(processed_files) == 1:
            processed_file = processed_files[0]
            st.download_button(
                label="Скачать итоговый XLSX-файл",
                data=processed_file["excel_bytes"],
                file_name=processed_file["output_file_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.download_button(
                label="Скачать ZIP-архив с итоговыми XLSX-файлами",
                data=zip_bytes,
                file_name="sroki_godnosti_results.zip",
                mime="application/zip",
            )


if __name__ == "__main__":
    main()
