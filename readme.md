# How to build

First, make sure that python and git is installed (https://www.python.org/ftp/python/3.13.3/python-3.13.3-amd64.exe, https://github.com/git-for-windows/git/releases/download/v2.49.0.windows.1/Git-2.49.0-64-bit.exe).

```
git clone https://github.com/ElizaLobanova/Parsing.git
pip install urllib3 beautifulsoup4 pandas openpyxl regex tqdm 
cd Parsing
```

Add files example.xlsx, urls_korting.txt, urls_hausedorf.txt. You can use folder "input_example" as an example.

# How to use

Just run `python parsing.py --help`, and read.