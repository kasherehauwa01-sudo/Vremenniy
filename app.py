"""Streamlit-приложение для преобразования Excel-файлов контроля сроков годности."""

from __future__ import annotations

import re
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
    E:P = месяцы. Строки заголовков и групп не попадут в результат, если
    в месячных столбцах нет годов вида 20XX.
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

        data = pd.read_excel(uploaded_file, header=None, engine=engine)
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

    if data.empty:
        raise FileProcessingError("Загруженный файл не содержит данных.")

    if data.shape[1] < REQUIRED_COLUMN_COUNT:
        raise FileProcessingError(
            "В файле должно быть минимум 16 столбцов: "
            "A = Код, B = Артикул, C = Наименование, D = Кол-во, E:P = месяцы."
        )

    return data


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
    logs: list[str] = [
        "Начата обработка данных.",
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


def main() -> None:
    """Запускает интерфейс Streamlit-приложения."""
    st.set_page_config(page_title="Преобразование сроков годности", layout="wide")

    st.title("Преобразование сроков годности")
    st.write(
        "Загрузите Excel-файл контроля сроков годности. Приложение обработает структуру "
        "со столбцами A = Код, B = Артикул, C = Наименование, D = Кол-во, E:P = месяцы "
        "и сформирует таблицу с колонками: Артикул, Количество, Срок годности до."
    )

    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл (.xls или .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=False,
    )

    if uploaded_file is None:
        st.info("Выберите файл Excel для обработки.")
        return

    if st.button("Обработать", type="primary"):
        try:
            with st.spinner("Идёт обработка файла..."):
                source_data = load_file(uploaded_file)
                result_data, logs = transform_data(source_data)
                excel_bytes = generate_excel(result_data)

            st.session_state["result_data"] = result_data
            st.session_state["excel_bytes"] = excel_bytes
            st.session_state["processing_logs"] = logs
            st.success(
                f"Файл успешно обработан. Сформировано строк: {len(result_data)}."
            )

            if result_data.empty:
                st.warning(
                    "В месячных столбцах E:P не найдено годов вида 20XX. "
                    "Итоговый файл будет содержать только заголовки."
                )
        except FileProcessingError as exc:
            st.session_state["processing_logs"] = [f"Ошибка: {exc}"]
            st.error(str(exc))
        except Exception as exc:
            st.session_state["processing_logs"] = [f"Непредвиденная ошибка: {exc}"]
            st.error(
                "Произошла непредвиденная ошибка при обработке файла. "
                f"Техническая информация: {exc}"
            )

    result_data = st.session_state.get("result_data")
    excel_bytes = st.session_state.get("excel_bytes")
    processing_logs = st.session_state.get("processing_logs")

    if processing_logs:
        st.subheader("Логи обработки")
        st.text_area(
            "Подробный журнал",
            value="\n".join(processing_logs),
            height=260,
            disabled=True,
        )

    if result_data is not None and excel_bytes is not None:
        st.subheader("Предпросмотр результата")
        st.dataframe(result_data.head(100), use_container_width=True)

        st.download_button(
            label="Скачать итоговый XLSX-файл",
            data=excel_bytes,
            file_name="sroki_godnosti_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
