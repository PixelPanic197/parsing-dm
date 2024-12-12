from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import os

# Путь к ChromeDriver
CHROMEDRIVER_PATH = "chromedriver.exe"

# Список артикулов для парсинга
product_ids = ['6623264']
# Подготовка списка для сохранения данных
data = []

# Настройка параметров Chrome
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Запуск браузера в безголовом режиме
chrome_options.add_argument("--disable-gpu")

# Инициализация драйвера Chrome
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    for product_id in product_ids:
        url = f"https://www.detmir.ru/product/index/id/{product_id}/"
        driver.get(url)

        try:
            # Ожидание загрузки страницы
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table tbody tr'))
            )

            # Извлечение данных
            product_code = driver.find_element(By.CSS_SELECTOR, 'table tbody tr:nth-child(1) td').text.strip() if driver.find_elements(By.CSS_SELECTOR, 'table tbody tr:nth-child(1) td') else "N/A"
            article_code = driver.find_element(By.CSS_SELECTOR, 'table tbody tr:nth-child(2) td').text.strip() if driver.find_elements(By.CSS_SELECTOR, 'table tbody tr:nth-child(2) td') else "N/A"
            price_with_promo = driver.find_element(By.CSS_SELECTOR, 'p.bvl.bvn[data-testid="price"]').text.strip() if driver.find_elements(By.CSS_SELECTOR, 'p.bvl.bvn[data-testid="price"]') else "N/A"
            price_without_promo = driver.find_element(By.CSS_SELECTOR, 'span.bvT').text.strip() if driver.find_elements(By.CSS_SELECTOR, 'span.bvT') else "N/A"
            promo_code_button = driver.find_element(By.CSS_SELECTOR, 'button.bvQ.bvS').text.strip() if driver.find_elements(By.CSS_SELECTOR, 'button.bvQ.bvS') else "N/A"

            # Извлечение данных по промокодам
            promo_blocks = driver.find_elements(By.CSS_SELECTOR, 'div.bw_1._5R.bw_9')

            smile = "N/A"
            snowflake = "N/A"
            winter = "N/A"
            fire = "N/A"

            for block in promo_blocks:
                promo_text = block.find_element(By.CSS_SELECTOR, 'p.bw_5').text.strip() if block.find_elements(By.CSS_SELECTOR, 'p.bw_5') else ""
                promo_name = block.find_element(By.CSS_SELECTOR, 'h4.bw_7').text.strip() if block.find_elements(By.CSS_SELECTOR, 'h4.bw_7') else ""
                promo_code = block.find_element(By.CSS_SELECTOR, 'span.bxc').text.strip() if block.find_elements(By.CSS_SELECTOR, 'span.bxc') else ""

                if promo_code == "УЛЫБКА":
                    smile = f"{promo_text} {promo_name}"
                elif promo_code == "СНЕЖОК":
                    snowflake = f"{promo_text} {promo_name}"
                elif promo_code == "ЗИМА":
                    winter = f"{promo_text} {promo_name}"
                elif promo_code == "ОГОНЕК":
                    fire = f"{promo_text} {promo_name}"

            # Добавление данных для каждого продукта
            data.append({
                'Код товара': product_code,
                'Артикул': article_code,
                'Цена с промокодом': price_with_promo,
                'Промокод': promo_code_button,
                'Цена без промокода': price_without_promo,
                'Улыбка': smile,
                'Цена Улыбка': "",
                'Снежок': snowflake,
                'Цена Снежок': "",
                'Зима': winter,
                'Цена Зима': "",
                'Огонек': fire,
                'Цена Огонек': ""
            })

            print(f"Данные успешно извлечены для артикула {product_id}.")

        except Exception as e:
            print(f"Ошибка при обработке артикула {product_id}: {e}")

finally:
    # Закрытие драйвера
    driver.quit()

# Проверка наличия файла Excel
file_path = 'product_data.xlsx'
if os.path.exists(file_path):
    # Загрузка существующих данных
    df_existing = pd.read_excel(file_path)
    df_new = pd.DataFrame(data)
    # Добавление новых данных
    df = pd.concat([df_existing, df_new], ignore_index=True)
else:
    # Создание нового DataFrame
    df = pd.DataFrame(data)

# Сохранение данных в Excel
df.to_excel(file_path, index=False)

print(f"Данные сохранены в файл '{file_path}'")