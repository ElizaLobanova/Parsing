from playwright.sync_api import sync_playwright
import os

def get_page_text(url, timeout=90000):
    with sync_playwright() as p:
        # Запускаем браузер с аргументами для большей совместимости
        browser = p.chromium.launch(
            headless=True,
            args=['--disable-blink-features=AutomationControlled']  # убираем автоматизацию
        )
        
        # Создаём контекст с реальным User-Agent и игнорированием SSL
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            ignore_https_errors=True,
            viewport={'width': 1920, 'height': 1080}  # размер окна как у реального браузера
        )
        
        page = context.new_page()
        page.set_default_timeout(timeout)  # устанавливаем глобальный тайм-аут
        
        try:
            # Переходим на страницу
            page.goto(url, wait_until='domcontentloaded')  # локальный тайм-аут уже не нужен
            # page.wait_for_load_state('networkidle')
            
            # Дополнительно можно подождать конкретный элемент, если известно, что страница динамическая
            # page.wait_for_selector('body', state='attached')
            
            text = page.inner_text('body')
            return text
        except Exception as e:
            print(f"Ошибка при обработке {url}: {e}")
            return None
        finally:
            browser.close()

content = get_page_text("https://www.franke.com/ru/ru/home-solutions/%D0%BF%D1%80%D0%BE%D0%B4%D1%83%D0%BA%D1%82%D1%8B/wine-coolers/product-detail-page.html/131.0632.993.html")
print(content)
with open('get_text_fact.txt', 'w', encoding='utf-8') as file:
    file.write(content)

