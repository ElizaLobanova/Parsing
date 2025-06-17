# How to build

First, make sure that python and git is installed (https://www.python.org/ftp/python/3.13.3/python-3.13.3-amd64.exe, https://github.com/git-for-windows/git/releases/download/v2.49.0.windows.1/Git-2.49.0-64-bit.exe).

```
git clone https://github.com/ElizaLobanova/Parsing.git
cd Parsing
pip install urllib3 beautifulsoup4 pandas openpyxl regex tqdm pyarrow ruwordnet pymorphy2
ruwordnet download
```

Add files example.xlsx, urls_korting.txt, urls_hausedorf.txt. You can use folder "input_example" as an example.

# How to use

Just run `python parsing.py --help`, `python synonyms_dict_update.py --help`, `python generate_syn_report.py --help` and `python compare.py --help`, and read. Example:
```
python parsing.py 34 True housedorf input_example/url_housedorf.txt input_example/example.xlsx result_housedorf.xlsx
python parsing.py 34 False korting input_example/url_korting.txt input_example/example.xlsx result.xlsx
python compare.py korting housedorf
python generate_syn_report.py housedorf syn_report_housedorf.xlsx
python generate_syn_report.py korting syn_report_korting.xlsx
```
If syn_report_housedorf.xlsx and syn_report_korting.xlsx are approved:
```
python synonyms_dict_update.py syn_report_housedorf.xlsx
python synonyms_dict_update.py syn_report_korting.xlsx
```