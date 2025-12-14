from ruwordnet import RuWordNet
import pymorphy2
import pandas as pd
from tqdm import tqdm
import re
from itertools import product
import argparse

synonyms_path='synonyms.txt'
all_characteristics_path='all_characteristics.xlsx'

# ---------------------------------------------------Остановить программу если не введено ни одного аргумента---------------------------------------------------
# Номер строки, взятый из аргументов запуска программы
parser = argparse.ArgumentParser(description="На основе .txt-файла, генерируемого программой парсинга, создаётся отчёт с предполагаемыми синонимами к " \
"незаписанным характеристикам. В любом случае будет выведено множество характеристик, для которых ещё не проводился поиск синонимов среди имеющихся в 1С")

parser.add_argument("site", type=str, help='Название типа сайта для парсинга. Возможные значения: korting, housedorf, dedietrich, falmec, vzug, asco, kuppersbush, konigin, evelux, franke, franke_dealer, elica, smeg, shaublorenz, shaublorenz_shop, graude, history, blanco, fashun, geizer, longran, makmart, aquaphor, mypremial, rivelato, topzero, ukinox, granfest, gerdamix.')
parser.add_argument("synonyms_report_path", type=str, default=None, help='Файл для записи отчёта о синонимах. Если не указан, то поиска синонимов не будет (учтите, что для более чем 5 характеристик составление отчёта занимает более 10 минут).')

args = parser.parse_args()
if not args.synonyms_report_path:
    raise ValueError("there's not enough arguments")

# ---------------------------------------------------------Добавить функции для сопоставления синонимов---------------------------------------------------------

wordnet = RuWordNet()
morph = pymorphy2.MorphAnalyzer()

def get_normal_form(word):
    return morph.parse(word)[0].normal_form

def are_synonyms(word1, word2):
    lemma1 = get_normal_form(word1)
    lemma2 = get_normal_form(word2)

    synsets1 = wordnet.get_synsets(lemma1)
    synsets2 = wordnet.get_synsets(lemma2)

    # Сравниваем наличие общих лемм в синсетах
    for s1 in synsets1:
        for s2 in synsets2:
            if s1.id == s2.id:
                return True
    return False

def list_synonyms_comparison(list1, list2):
    return [are_synonyms(word1, word2) for word1, word2 in zip(list1, list2)]

# -----------------------------------Добавить функцию для создания отчёта о сопоставлении колонок тем, что уже существуют в 1С-----------------------------------

def parse_custom_dict_line(line):
    """
    Разбирает строку из словаря: <характеристика>: <синоним1>; <синоним2>; ...| <антисиноним1>, <антисиноним2>, ...
    """
    base, *rest = line.strip().split(':')
    if not rest:
        return base.strip(), set(), set()
    syn_ant = rest[0].split('|')
    synonyms = set(map(str.strip, syn_ant[0].split(';'))) if syn_ant[0] else set()
    antonyms = set(map(str.strip, syn_ant[1].split(','))) if len(syn_ant) > 1 else set()
    return base.strip(), synonyms, antonyms

def load_existing_synonyms(file_path):
    syn_dict = {}
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            base, synonyms, antonyms = parse_custom_dict_line(line)
            syn_dict[base] = {'synonyms': synonyms, 'antonyms': antonyms}
    return syn_dict

def tokenize(text):
    return set(re.findall(r'\w+', text.lower()))

def are_words_possibly_synonyms(words1, words2, are_synonims_func):
    pairs = product(words1, words2)
    for w1, w2 in pairs:
        if w1 == w2:
            return True
        if are_synonims_func(w1, w2):
            return True
    return False


def generate_synonym_report(existing_dict_path, all_1c_chars_path, new_chars, are_synonims_func, output_excel_path):
    custom_dict = load_existing_synonyms(existing_dict_path)
    all_chars = pd.read_excel(all_1c_chars_path, header=None).iloc[0].astype(str).tolist()
    result_rows = []

    for c1 in tqdm(all_chars):
        for c2 in new_chars:
            if c1 == c2:
                continue
            tokens1 = tokenize(c1)
            tokens2 = tokenize(c2)
            if are_words_possibly_synonyms(tokens1, tokens2, are_synonims_func):
                result_rows.append((c1, c2, None))  # None для ручной отметки

    df_result = pd.DataFrame(result_rows, columns=["base_char", "compared_char", "label"])
    df_result.to_excel(output_excel_path, index=False)
    
# -----------------------------------------------------------------------Запуск генерации-----------------------------------------------------------------------

with open(f'unaccepted_syn_{args.site}.txt', 'r', encoding='utf-8') as f:
    data = f.read()

unsyn_list = list(map(str, data.strip().split('; ')))

generate_synonym_report(
    synonyms_path,
    all_characteristics_path,
    unsyn_list,
    are_synonyms,
    args.synonyms_report_path
)