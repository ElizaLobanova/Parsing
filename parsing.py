import urllib3
from bs4 import BeautifulSoup, NavigableString
import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment
import re
from tqdm import tqdm
import warnings
warnings.filterwarnings("ignore")
import argparse
import os

# ---------------------------------------------------Остановить программу если не введено ни одного аргумента---------------------------------------------------
# Номер строки, взятый из аргументов запуска программы
parser = argparse.ArgumentParser(description="Парсинг данных в excel-файл с одного типа сайтов. В качестве входного файла используется выгрузка из 1С. " \
"Выходной файл создаётся по образу и подобию входного, является результатом парсинга. Если какая-то ячейка уже была заполнена в excel-файле, то она не " \
"будет перезаписана. Возможно дополнение файла столбцами на основе найденных на сайте характеристик. Дополнительно сохраняется вспомогательный csv-файл " \
"для дальнейшего сравнения данных с разных сайтов")

parser.add_argument("start_row", type=int, help="Номер строки в excel, начиная с которой необходимо писать данные")
parser.add_argument("append", type=str, help="Добавлять ли в конец дополнительные столбцы с незаписанными данными сайтов. Возможные значения: True, " \
"False")
parser.add_argument("site", type=str, help='Название типа сайта для парсинга. Возможные значения: korting, housedorf')
parser.add_argument("urls_source", type=str, help='Файл со ссылками на карточки товаров. Порядок должен соответствовать расположению наименований ' \
'номенклатуры в excel-файле, указанном как входной')
parser.add_argument("input_path", type=str, help='Путь к входному excel-файлу')
parser.add_argument("output_path", type=str, help='Путь к выходному excel-файлу')

args = parser.parse_args()
if not args.output_path:
    raise ValueError("there's not enough arguments")
args.append = True if args.append.lower() == 'true' else False

# ------------------------------------------------------------------------Спарсить данные------------------------------------------------------------------------

# Функции для парсинга
def parse_korting_page(html_code):
    soup = BeautifulSoup(html_code, 'html.parser')
    tabs_lists = soup.find_all('ul', class_='tabs-settings__list')
    data = {}

    for ul in tabs_lists:
        for li in ul.find_all('li'):
            text = li.get_text(strip=True, separator="; ")
            split_text = text.split(":;", 1)
            if len(split_text) == 2:
                key, value = split_text
                data[key.strip()] = value.strip()
            else:
                data[split_text[0].strip()] = ""

    return data

def extract_first_visible_text(tag):
    for desc in tag.descendants:
        if isinstance(desc, str):  # Это NavigableString
            text = desc.strip()
            if text:
                return text
    return None

def clean_value_div(value_div):
    # 1. Удалить все <span>
    for span in value_div.find_all("span"):
        span.decompose()

    # 2. Разделить по <br> — создаём список на основе HTML с разделителем
    parts = str(value_div).split('<br')

    values = []

    for part_html in parts:
        # Восстанавливаем HTML-тег <br>, если он был отрезан
        if not part_html.startswith('>'):
            part_html = '<br' + part_html

        part_soup = BeautifulSoup(part_html, 'html.parser')

        # 3. Найти первый видимый текст
        for desc in part_soup.descendants:
            if isinstance(desc, NavigableString):
                text = desc.strip()
                if text:
                    values.append(text)
                    break  # только первое вхождение

    # 4. Склеить с разделителем "; "
    return "; ".join(values)

def parse_hausedorf_page(html_code):
    soup = BeautifulSoup(html_code, 'html.parser')
    fields = soup.find_all('div', class_='detail-properties__field')
    data = {}

    for field in fields:
        name_div = field.find('div', class_='detail-properties__name')
        value_div = field.find('div', class_='detail-properties__value')

        if name_div and value_div:
            key = extract_first_visible_text(name_div)
            value = clean_value_div(value_div)

            if key and value:
                data[re.sub(r'\s+', ' ', key).strip()] = value.replace(">\n", "")

    return data


def create_src(file_path, parser_func):
    """
    Универсальный загрузчик таблицы характеристик с разных сайтов.

    :param file_path: путь к файлу с URL (один URL на строку)
    :param parser_func: функция, которая получает HTML-код и возвращает словарь {ключ: значение}
    :return: DataFrame с объединёнными результатами
    """
    http = urllib3.PoolManager()
    df_all = pd.DataFrame()

    with open(file_path, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f]

    for url in tqdm(urls):
        if url:
            try:
                response = http.request('GET', url)
                html_code = response.data.decode()
                data = parser_func(html_code) 

                if not isinstance(data, dict):
                    raise ValueError("parser_func должна возвращать словарь!")

                row_df = pd.DataFrame([data])
                df_all = pd.concat([df_all, row_df], ignore_index=True)

            except Exception as e:
                print(f"Ошибка при обработке {url}: {e}")
        else:
            df_all.loc[len(df_all)] = None

    return df_all.where(pd.notnull(df_all), None)

if args.site == 'korting':
    df_src = create_src(args.urls_source, parse_korting_page)
elif args.site == 'housedorf':
    df_src = create_src(args.urls_source, parse_hausedorf_page)
else:
    raise ValueError("There're no parse function for this site")

# ---------------------------------------------Записать и вернуть дополненный названиями номенклатуры DataFrame---------------------------------------------

def write_dest(ref_file_path, result_file_path, df_src, start_row_index):
    # Путь к файлу Excel
    wb = load_workbook(ref_file_path)
    ws = wb.active  # или wb['SheetName']

    # Поля файла-приёмника
    row_header = [cell.value for cell in ws[1]]

    # Сопоставление колонок
    src_cols_lower = {col.lower(): col for col in df_src.columns}
    ws_cols_lower = {i: str(header).strip().lower() if header else "" for i, header in enumerate(row_header)}
    matched_columns = []
    common_cols = set()
    missing_cols = set()
    nomenclature_col_idx = None

    for col_idx, header_lower in ws_cols_lower.items():
        if header_lower == "номенклатура":
            nomenclature_col_idx = col_idx
        if header_lower in src_cols_lower:
            matched_columns.append((col_idx, src_cols_lower[header_lower]))
            common_cols.add(src_cols_lower[header_lower])
        else:
            missing_cols.add(header_lower)

    if nomenclature_col_idx is None:
        raise ValueError("Колонка 'Номенклатура' не найдена в файле-приёмнике")

    # Считываем значения "Номенклатура"
    nomenclature_values = []
    for i in range(len(df_src)):
        cell_value = ws.cell(row=start_row_index + i, column=nomenclature_col_idx + 1).value
        nomenclature_values.append(cell_value)

    # Запись данных
    for i, (_, row_src) in enumerate(df_src.iterrows()):
        for col_idx, src_col in matched_columns:
            cell = ws.cell(row=start_row_index + i, column=col_idx + 1)
            if cell.value in [None, ""]:
                cell.value = row_src[src_col]

    # Сохранение
    wb.save(result_file_path)

    # Подготовка выходного DataFrame
    result_df = df_src[list(common_cols)].copy()
    result_df.insert(0, "Номенклатура", pd.Series(nomenclature_values))

    return result_df, missing_cols

resultdf, _ = write_dest(args.input_path, args.output_path, df_src, args.start_row)
resultdf.to_parquet(f"{args.site}_auxiliary.parquet")

# -------------------------------------------Сохранить незаписанные данные в дополнительные колонки или отдельный файл-------------------------------------------

def append_dataframe_to_excel(df: pd.DataFrame, file_path: str, result_path: str, start_row: int):
    # Проверка, существует ли файл
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
    else:
        # Если файла нет, создаем новый
        wb = Workbook()
    
    ws = wb.active

    # Найдём первую пустую ячейку в первой строке
    col_index = 1
    while ws.cell(row=1, column=col_index).value is not None:
        col_index += 1

    # Записываем названия колонок DataFrame в первую строку, начиная с найденной колонки
    for i, col_name in enumerate(df.columns):
        ws.cell(row=1, column=col_index + i, value=col_name)

    # Автонастройка ширины столбцов по первой строке
    for col_idx, cell in enumerate(ws[1], start=col_index):
        max_length = len(str(cell.value)) if cell.value else 0
        col_letter = cell.column_letter
        ws.column_dimensions[col_letter].width = max_length + 2  # +2 для отступа

    # Применение стилей и переносов
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')  # включаем перенос текста

    # Записываем данные DataFrame начиная со start_row
    for row_offset, row in enumerate(dataframe_to_rows(df, index=False, header=False)):
        for i, value in enumerate(row):
            ws.cell(row=start_row + row_offset, column=col_index + i, value=value)

    # Сохраняем файл
    wb.save(result_path)

def save_missing(df1, filepath):

    # Создаём ExcelWriter
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        row = 0

        # Запись df1
        df1.to_excel(writer, index=False, startrow=row)
        row += len(df1) + 2  # +1 за заголовок, +1 за пустую строку

        # Автоматическая установка ширины колонок
        worksheet = writer.sheets['Sheet1']
        for column_cells in worksheet.columns:
            max_length = 0
            column = column_cells[0].column
            for cell in column_cells:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2
            worksheet.column_dimensions[get_column_letter(column)].width = adjusted_width

comparison_cols = [col for col in resultdf.columns if col != 'Номенклатура']
missingdf = df_src.drop(columns=comparison_cols)
if not args.append:
    save_missing(missingdf, f'missing_{args.site}.xlsx')
else:
    append_dataframe_to_excel(missingdf, args.output_path, args.output_path, args.start_row)

print("Successfully finished")