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

# В исходном файле месяцы расположены в столбцах C:N.
# Индексы pandas начинаются с 0, поэтому C = 2, D = 3, ..., N = 13.
MONTH_COLUMNS = {
    2: "01",
    3: "02",
    4: "03",
    5: "04",
    6: "05",
    7: "06",
    8: "07",
    9: "08",
    10: "09",
    11: "10",
    12: "11",
    13: "12",
}

RESULT_COLUMNS = ["Артикул", "Количество", "Срок годности до"]


class FileProcessingError(Exception):
    """Ошибка, понятная пользователю при чтении или обработке файла."""



def load_file(uploaded_file: BinaryIO) -> pd.DataFrame:
    """
    Загружает Excel-файл в DataFrame.

    Для максимальной совместимости файл читается без строки заголовков
    (header=None), потому что структура задана позициями столбцов:
    A = артикул, B = количество, C:N = месяцы.
    Если в файле есть строка заголовков, она не попадёт в результат,
    так как в ней обычно нет годов вида 20XX в месячных столбцах.
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

    if data.shape[1] < 14:
        raise FileProcessingError(
            "В файле должно быть минимум 14 столбцов: A = Артикул, B = Количество, C:N = месяцы."
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



def transform_data(data: pd.DataFrame) -> pd.DataFrame:
    """
    Преобразует исходную таблицу в таблицу со сроками годности.

    Для каждой строки исходного файла просматриваются месячные столбцы C:N.
    Каждый найденный год вида 20XX превращается в отдельную строку результата
    с датой в строгом формате ММ.ГГГГ.
    """
    if data.shape[1] < 14:
        raise FileProcessingError(
            "Недостаточно столбцов для обработки. Ожидаются столбцы A:N."
        )

    result_rows: list[dict[str, object]] = []

    # Сортировка не выполняется: порядок обхода повторяет исходный порядок строк,
    # внутри строки — порядок месяцев C:N, внутри ячейки — порядок найденных годов.
    for _, row in data.iterrows():
        article = row.iloc[0]
        quantity = _normalize_quantity(row.iloc[1])

        for column_index, month in MONTH_COLUMNS.items():
            cell_text = _cell_to_text(row.iloc[column_index])
            years = YEAR_PATTERN.findall(cell_text)

            for year in years:
                result_rows.append(
                    {
                        "Артикул": article,
                        "Количество": quantity,
                        "Срок годности до": f"{month}.{year}",
                    }
                )

    return pd.DataFrame(result_rows, columns=RESULT_COLUMNS)



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
        "Загрузите Excel-файл контроля сроков годности. Приложение найдёт годы "
        "в столбцах C:N и сформирует таблицу с колонками: Артикул, Количество, Срок годности до."
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
                result_data = transform_data(source_data)
                excel_bytes = generate_excel(result_data)

            st.session_state["result_data"] = result_data
            st.session_state["excel_bytes"] = excel_bytes
            st.success(
                f"Файл успешно обработан. Сформировано строк: {len(result_data)}."
            )

            if result_data.empty:
                st.warning(
                    "В месячных столбцах C:N не найдено годов вида 20XX. "
                    "Итоговый файл будет содержать только заголовки."
                )
        except FileProcessingError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(
                "Произошла непредвиденная ошибка при обработке файла. "
                f"Техническая информация: {exc}"
            )

    result_data = st.session_state.get("result_data")
    excel_bytes = st.session_state.get("excel_bytes")

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
