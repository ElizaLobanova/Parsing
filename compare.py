import urllib3
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
import re
from tqdm import tqdm
import warnings
warnings.filterwarnings("ignore")
import argparse
import os

# ---------------------------------------------------Остановить программу если не введено ни одного аргумента---------------------------------------------------
# Номер строки, взятый из аргументов запуска программы
parser = argparse.ArgumentParser(description="Сравнивает результаты парсинга с двух сайтов (убедитесь, что парсинг с каждого из указываемых в аргументах " \
"к этой программе сайтов проведён с помощью программы parsing.py). Программа использует сохранённые программой парсинга вспомогательные файлы. " \
"Результат сравнения записывается в excel-файл с названием, начинающемся с 'comparison', с упоминанием сайтов, по которым шло сравнение.")

parser.add_argument("site1", type=str, help="Тип первого сайта, участвующего в сравнении")
parser.add_argument("site2", type=str, help='Тип второго сайта, участвующего в сравнении')

args = parser.parse_args()
if not args.site2:
    raise ValueError("there's not enough arguments")
if args.site1 == args.site2:
    raise ValueError("the arguments coincide, comparison is impossible")

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

resultdf_1 = pd.read_parquet(f"{args.site1}_auxiliary.parquet")
resultdf_2 = pd.read_parquet(f"{args.site2}_auxiliary.parquet")
common_cols = resultdf_1.columns.intersection(resultdf_2.columns)
comp_result = compare_dataframes(resultdf_1, resultdf_2, args.site1, args.site2)

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

save_comparison_to_excel(comp_result, f'comparison_{args.site1}_vs_{args.site2}.xlsx')