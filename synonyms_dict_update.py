import pandas as pd
from collections import defaultdict
import os
import argparse
synonyms_path='synonyms.txt'

# ---------------------------------------------------Остановить программу если не введено ни одного аргумента---------------------------------------------------
# Номер строки, взятый из аргументов запуска программы
parser = argparse.ArgumentParser(description="Обновляет словарь синонимов, который используется в программе parsing.py. ")

parser.add_argument("report_path", type=str, help="Путь к excel-файлу с утверждённым отчётом о синонимах, который был сгенерирован программой parsing.py")

args = parser.parse_args()
if not args.report_path:
    raise ValueError("there's not enough arguments")

# --------------------------------------------------------------Обновить словарь синонимов из Excel--------------------------------------------------------------
def load_synonym_dict(dict_path):
    syn_dict = defaultdict(lambda: {"synonyms": set(), "antisynonyms": set()})
    if not os.path.exists(dict_path):
        return syn_dict
    
    with open(dict_path, "r", encoding="utf-8") as f:
        for line in f:
            if ":" not in line:
                continue
            key, rest = line.strip().split(":", 1)
            syn_part, *anti_part = rest.strip().split("|")
            syns = set(map(str.strip, syn_part.strip().split(";"))) if syn_part.strip() else set()
            antis = set(map(str.strip, anti_part[0].strip().split(","))) if anti_part and anti_part[0].strip() else set()
            syn_dict[key.strip()]["synonyms"].update(syns)
            syn_dict[key.strip()]["antisynonyms"].update(antis)
    return syn_dict


def save_synonym_dict(syn_dict, dict_path):
    with open(dict_path, "w", encoding="utf-8") as f:
        for key in sorted(syn_dict.keys()):
            syns = "; ".join(sorted(syn_dict[key]["synonyms"]))
            antis = ", ".join(sorted(syn_dict[key]["antisynonyms"]))
            line = f"{key}: {syns} | {antis}\n"
            f.write(line)


def update_synonym_dict_from_excel(excel_path, dict_path):
    df = pd.read_excel(excel_path)
    if not {"base_char", "compared_char", "label"}.issubset(df.columns):
        raise ValueError("Excel должен содержать столбцы: base_char, compared_char, label")
    
    syn_dict = load_synonym_dict(dict_path)

    for _, row in df.iterrows():
        base = row["base_char"].strip()
        comp = row["compared_char"].strip()
        label = row["label"]

        if label == 1:
            syn_dict[base]["synonyms"].add(comp)
        elif label == 0.5:
            syn_dict[base]["synonyms"].add("*"+comp)
        elif label == 0 or pd.isna(label):
            syn_dict[base]["antisynonyms"].add(comp)
        else:
            continue  # Пропустить некорректные значения

    # Удалить пересекающиеся значения
    for base in syn_dict:
        overlap1 = syn_dict[base]["synonyms"] & syn_dict[base]["antisynonyms"]
        overlap2 = set(map(lambda x: x.split("*")[1] if len(x.split("*")) > 1 else x, syn_dict[base]["synonyms"])) & syn_dict[base]["antisynonyms"]
        syn_dict[base]["antisynonyms"] -= overlap1
        syn_dict[base]["antisynonyms"] -= overlap2

    save_synonym_dict(syn_dict, dict_path)
    print(f"Обновлённый словарь сохранён в {dict_path}")

update_synonym_dict_from_excel(args.report_path, synonyms_path)