import urllib3
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
import re
from tqdm import tqdm
import argparse
import warnings
warnings.filterwarnings("ignore")

# ---------------------------------------------------Остановить программу если не введено ни одного аргумента---------------------------------------------------
# Номер строки, взятый из аргументов запуска программы
parser = argparse.ArgumentParser(description="")
parser.add_argument("str_index", type=int, help="Номер строки в excel, начиная с которой необходимо писать данные")
args = parser.parse_args()

# ------------------------------------------------------------Протестировать парсинг со второго сайта------------------------------------------------------------

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

def parse_hausedorf_page(html_code):
    soup = BeautifulSoup(html_code, 'html.parser')
    fields = soup.find_all('div', class_='detail-properties__field')
    data = {}

    for field in fields:
        name_div = field.find('div', class_='detail-properties__name')
        value_div = field.find('div', class_='detail-properties__value')

        if name_div and value_div:
            raw_key = name_div.find(text=True, recursive=False)
            value = value_div.find(text=True, recursive=False)
            if raw_key and value:
                key = re.sub(r'\s+', ' ', raw_key).strip()
                data[key] = value

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

df_src_korting = create_src('urls_korting.txt', parse_korting_page)
df_src_housedorf = create_src('urls_hausedorf.txt', parse_hausedorf_page)
while len(df_src_korting) > len(df_src_housedorf):
    df_src_housedorf.loc[len(df_src_housedorf)] = None
while len(df_src_korting) < len(df_src_housedorf):
    df_src_korting.loc[len(df_src_korting)] = None

# -------------------------------------------По итогу заполнения вернуть дополненный названиями номенклатуры DataFrame-------------------------------------------

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

resultdf_korting, missing_korting = write_dest('example.xlsx', 'result_korting.xlsx', df_src_korting, args.str_index)
resultdf_housedorf, missing_housedorf = write_dest('example.xlsx', 'result_housedorf.xlsx', df_src_housedorf, args.str_index)

# Записать korting поверх housdorf'а
write_dest('result_housedorf.xlsx', 'result.xlsx', df_src_korting, 34)

# -------------------------------------------------------------------Сравнить записанные данные-------------------------------------------------------------------

def compare_dataframes(df1: pd.DataFrame, df2: pd.DataFrame, name1: str = 'df1', name2: str = 'df2') -> pd.DataFrame:
    """
    Сравнивает два DataFrame по общим столбцам, включая "Номенклатура" как ключ.
    Добавляет префиксы к столбцам (кроме "Номенклатура") и формирует колонку diff_columns.
    """
    if 'Номенклатура' not in df1.columns or 'Номенклатура' not in df2.columns:
        raise ValueError("Оба DataFrame должны содержать колонку 'Номенклатура'")

    # Определим общие характеристики (кроме "Номенклатура")
    common_columns = df1.columns.intersection(df2.columns).difference(['Номенклатура'])

    # Сузим входные DataFrame'ы до нужных колонок
    df1_reduced = df1[['Номенклатура'] + list(common_columns)].copy()
    df2_reduced = df2[['Номенклатура'] + list(common_columns)].copy()

    # Переименуем характеристики с префиксами, "Номенклатура" оставим без изменений
    df1_renamed = df1_reduced.rename(columns={col: f'{name1}_{col}' for col in common_columns})
    df2_renamed = df2_reduced.rename(columns={col: f'{name2}_{col}' for col in common_columns})

    # Объединение по "Номенклатура"
    df_merged = pd.merge(df1_renamed, df2_renamed, on='Номенклатура', how='outer')

    # Функция для сравнения значений по строке
    def get_differences(row):
        diffs = []
        for col in common_columns:
            val1 = row.get(f'{name1}_{col}', None)
            val2 = row.get(f'{name2}_{col}', None)
            if pd.isna(val1) or pd.isna(val2):
                continue
            if val1 != val2:
                diffs.append(col)
        return ', '.join(diffs)

    # Добавление столбца с различиями
    df_merged['diff_columns'] = df_merged.apply(get_differences, axis=1)

    return df_merged

common_cols = resultdf_korting.columns.intersection(resultdf_housedorf.columns)
comp_result = compare_dataframes(resultdf_korting[list(common_cols)], resultdf_housedorf[list(common_cols)], 'Korting', 'Hausedorf')

# --------------------------------------------Сохранить незаписанные данные в excel и записать общие пропущенные колонки--------------------------------------------

# Незаписанные данные
comparison_cols_korting = [col for col in resultdf_korting.columns if col != 'Номенклатура']
comparison_cols_housedorf = [col for col in resultdf_housedorf.columns if col != 'Номенклатура']
missingdf_korting = df_src_korting.drop(columns=comparison_cols_korting)
missingdf_housedorf = df_src_housedorf.drop(columns=comparison_cols_housedorf)
missingdf_korting.insert(0, 'Номенклатура', resultdf_korting['Номенклатура'].values)
missingdf_housedorf.insert(0, 'Номенклатура', resultdf_housedorf['Номенклатура'].values)

# Пропущенные колонки
common_missing = missing_korting.intersection(missing_housedorf)

# Сохранение
def save_missing(df1, df2, my_set, filepath):
    # Преобразуем set в строку (одна строка, несколько колонок)
    df_set_row = pd.DataFrame([list(my_set)])

    # Создаём ExcelWriter
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        row = 0

        # Запись df1
        df1.to_excel(writer, index=False, startrow=row)
        row += len(df1) + 2  # +1 за заголовок, +1 за пустую строку

        # Запись df2
        df2.to_excel(writer, index=False, startrow=row)
        row += len(df2) + 2

        # Запись множества в строку, без подписи и индексов
        df_set_row.to_excel(writer, index=False, header=False, startrow=row)

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

save_missing(missingdf_korting, missingdf_housedorf, common_missing, 'missing.xlsx')

# --------------------------------------------------------------Сохранить результат сравнения в excel--------------------------------------------------------------

def save_comparison_to_excel(df: pd.DataFrame, filename: str):
    # Сохранить в Excel без форматирования сначала
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comparison')
    
    # Открыть книгу с openpyxl
    wb = load_workbook(filename)
    ws = wb['Comparison']

    # Настройка ширины колонок
    for col_idx, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        max_length = max((len(str(cell.value)) if cell.value else 0) for cell in col_cells)
        adjusted_width = max_length + 2
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    # Выделим разницу цветом
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    headers = [cell.value for cell in ws[1]]
    col_indexes = {col: idx + 1 for idx, col in enumerate(headers)}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        diff_text = row[col_indexes['diff_columns'] - 1].value
        if diff_text:
            differing_cols = [col.strip() for col in diff_text.split(',')]
            for col in differing_cols:
                # Попробуем подсветить оба варианта значений
                for prefix in ['Korting_', 'Hausedorf_']:
                    full_col = prefix + col
                    if full_col in col_indexes:
                        cell = row[col_indexes[full_col] - 1]
                        cell.fill = yellow_fill
            # diff_columns — красной заливкой
            row[col_indexes['diff_columns'] - 1].fill = red_fill

    # Добавим автофильтр
    ws.auto_filter.ref = ws.dimensions

    wb.save(filename)

save_comparison_to_excel(comp_result, 'comparison.xlsx')

print("Successfully finished")