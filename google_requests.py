from googleapiclient.discovery import build
import json

API_KEY = "AIzaSyCrx5HLjkeDcg_7v799xd6l9o2tb--XcXw"
CX_ID = 'f237459216895469f'
QUERY = "всё о животных"

def google_search(query, api_key, cx_id):
    # Создание объекта сервиса для Custom Search API v1
    service = build("customsearch", "v1", developerKey=api_key)
    
    # Выполнение запроса
    res = service.cse().list(
        q=query,
        cx=cx_id,
        num=5 # Количество результатов
    ).execute()
    
    return res

results = google_search(QUERY, API_KEY, CX_ID)

# Вывод результатов
if 'items' in results:
    for i, item in enumerate(results['items'], 1):
        print(f"Результат {i}:")
        print(f"Заголовок: {item['title']}")
        print(f"Ссылка: {'bookyourhunt' in item['link']}")
        print(f"Краткое описание: {item['snippet']}\n")
else:
    print("Результаты поиска не найдены.")
