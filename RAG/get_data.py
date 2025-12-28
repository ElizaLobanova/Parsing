from googleapiclient.discovery import build
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.keys import Keys
import pyperclip

API_KEY = "AIzaSyCrx5HLjkeDcg_7v799xd6l9o2tb--XcXw" 
CX_ID = 'f237459216895469f'
# QUERY = "всё о животных"

def load_articules(path: str) -> list[str]:
    articles = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split(" - ", 1)
            if len(parts) == 2:
                articles.append(parts[1].strip())
            else:
                # если строки без разделителя — игнорируем
                continue
    return articles


def google_search(query, api_key, cx_id):
    service = build("customsearch", "v1", developerKey=api_key)
    res = service.cse().list(
        q=query,
        cx=cx_id,
        num=3
    ).execute()
    return res

def get_page_text_ctrl(driver):
    body = driver.find_element("tag name", "body")
    body.send_keys(Keys.CONTROL, 'a')
    body.send_keys(Keys.CONTROL, 'c')

    # Забрать из буфера обмена
    return pyperclip.paste()

# --- Основной код ---
for QUERY in load_articules("Ненайденные/Ненайденные.txt"):
    time.sleep(10)
    results = google_search(QUERY, API_KEY, CX_ID)

    if 'items' in results:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        for i, item in enumerate(results['items'], 1):
            url = item["link"]
            print(f"[{i}] Загружаю: {url}")

            try:
                driver.get(url)
                time.sleep(3)
                text = get_page_text_ctrl(driver)

                with open(f"rag_data/page_{i}_{QUERY}.txt", "w", encoding="utf-8") as f:
                    f.write(text)
                print(f"Сохранено в page_{i}_{QUERY}.txt")
            except Exception as e:
                print(f"Ошибка при обработке {url}: {e}")

        driver.quit()
    else:
        print("Результаты поиска не найдены.")
