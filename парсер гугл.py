from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from bs4 import BeautifulSoup
import pandas as pd

# Инициализация веб-драйвера
driver = webdriver.Chrome() 
# Гугл карты
# url = 'https://www.google.com/maps/place/CDEK/@55.9063864,37.2555477,17z/data=!4m8!3m7!1s0x46b5417929804f31:0x7b4e2a789ee0bcf9!8m2!3d55.9063864!4d37.2555477!9m1!1b1!16s%2Fg%2F11fvqnddkp?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D'  # Замените на свой URL
url = 'https://www.google.com/maps/place/CDEK/@55.843726,37.35424,17z/data=!4m8!3m7!1s0x46b547e601912a11:0x6d309bfcfd079430!8m2!3d55.843726!4d37.35424!9m1!1b1!16s%2Fg%2F11qn9bfd8n?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #MTN69
# url = 'https://www.google.com/maps/place/CDEK/@55.71561,37.892436,17z/data=!4m8!3m7!1s0x414ac95379a766ef:0xe5ebe1601c776be5!8m2!3d55.71561!4d37.892436!9m1!1b1!16s%2Fg%2F11nfvh6chw?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #MSK2327
# url = 'https://www.google.com/maps/place/CDEK/@55.538354,37.537338,17z/data=!4m8!3m7!1s0x414aadf0ff5072cb:0x16aa8c97be2372f!8m2!3d55.538354!4d37.537338!9m1!1b1!16s%2Fg%2F11pwwybw5r?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #BTV65
# url = 'https://www.google.com/maps/place/CDEK/@55.8444723,37.2685096,12z/data=!4m8!3m7!1s0x46b547f8bd6705c9:0xa8f347a5492c61b!8m2!3d55.834376!4d37.277078!9m1!1b1!16s%2Fg%2F11rfm08h3d?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #KRN10
driver.get(url)
input("Пройдите капчу и нажмите Enter, чтобы продолжить...")
# Функция скроллинга страницы
def scroll():
    elements = driver.find_elements(By.CLASS_NAME, 'TFQHme ')
    count_el = len(elements)
    count_el2 = 0
    while count_el != count_el2:
        count_el = len(elements)
        last_element = elements[-1]
        driver.execute_script("arguments[0].scrollIntoView(true);", last_element)
        sleep(4)  # Даем время для загрузки новых элементов
        elements = driver.find_elements(By.CLASS_NAME, 'TFQHme ')
        count_el2 = len(elements)
    return count_el2

# Выполняем скроллинг
scroll()
print("Скроллинг завершен!")

# Получаем HTML код страницы
data = driver.page_source

# Разбираем HTML с помощью BeautifulSoup
soup = BeautifulSoup(data, 'html.parser')

# Находим все блоки отзывов
review_blocks = soup.find_all(class_='jJc9Ad')

# Инициализация структуры данных
keys = {'№': [], 'Имя': [], 'Рейтинг': [], 'Комментарий': []}

# Парсим данные из каждого блока
for idx, block in enumerate(review_blocks, start=1):
    keys['№'].append(idx)  # Номер отзыва

    # Имя комментатора
    name = block.find('div', {'class': 'd4r55'})
    keys['Имя'].append(name.text.strip() if name else 'Ошибка')

    # Рейтинг
    try:
        # Ищем элемент с классом 'kvMYJc' и атрибутом 'role="img"'
        rating_tag = block.find('span', {'class': 'kvMYJc', 'role': 'img'})
        
        # Проверяем, есть ли атрибут 'aria-label', и извлекаем рейтинг
        if rating_tag and 'aria-label' in rating_tag.attrs:
            rating_text = rating_tag['aria-label']  # Пример: "1 star"
            rating_value = rating_text.split()[0]  # Извлекаем только число
            keys['Рейтинг'].append(rating_value)
        else:
            keys['Рейтинг'].append('Ошибка')
    except Exception as e:
        print(f"Ошибка при обработке рейтинга: {e}")
        keys['Рейтинг'].append('Ошибка')

    # Комментарий
    try:
        # Ищем элемент <span> с классом 'wiI7pd'
        comment_tag = block.find('span', {'class': 'wiI7pd'})
        
        # Проверяем, найден ли элемент, и извлекаем текст
        if comment_tag:
            comment_text = comment_tag.text.strip()  # Убираем лишние пробелы
            keys['Комментарий'].append(comment_text)
        else:
            keys['Комментарий'].append('Ошибка')
    except Exception as e:
        print(f"Ошибка при обработке текста отзыва: {e}")
        keys['Комментарий'].append('Ошибка')

# Закрываем браузер
driver.quit()

# Проверяем длины всех списков и синхронизируем их
max_length = max(len(keys[col]) for col in keys)
for col in keys:
    while len(keys[col]) < max_length:
        keys[col].append('Ошибка')

# Создаем DataFrame и сохраняем в Excel
df = pd.DataFrame(keys)
df.to_excel('data.xlsx', index=False)
print("Все отзывы успешно сохранены!")
