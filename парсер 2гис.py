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
# 2гис
url = 'https://2gis.ru/moscow/search/сдэк%20дубравная/firm/70000001042746853/37.354428%2C55.843643/tab/reviews?m=37.357135%2C55.84591%2F16.6'  # YURL6
# url = 'https://www.google.com/maps/place/CDEK/@55.843726,37.35424,17z/data=!4m8!3m7!1s0x46b547e601912a11:0x6d309bfcfd079430!8m2!3d55.843726!4d37.35424!9m1!1b1!16s%2Fg%2F11qn9bfd8n?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #MTN69
# url = 'https://www.google.com/maps/place/CDEK/@55.71561,37.892436,17z/data=!4m8!3m7!1s0x414ac95379a766ef:0xe5ebe1601c776be5!8m2!3d55.71561!4d37.892436!9m1!1b1!16s%2Fg%2F11nfvh6chw?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #MSK2327
# url = 'https://www.google.com/maps/place/CDEK/@55.538354,37.537338,17z/data=!4m8!3m7!1s0x414aadf0ff5072cb:0x16aa8c97be2372f!8m2!3d55.538354!4d37.537338!9m1!1b1!16s%2Fg%2F11pwwybw5r?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #BTV65
# url = 'https://www.google.com/maps/place/CDEK/@55.8444723,37.2685096,12z/data=!4m8!3m7!1s0x46b547f8bd6705c9:0xa8f347a5492c61b!8m2!3d55.834376!4d37.277078!9m1!1b1!16s%2Fg%2F11rfm08h3d?entry=ttu&g_ep=EgoyMDI1MDExMC4wIKXMDSoASAFQAw%3D%3D' #KRN10
driver.get(url)
# Функция скроллинга страницы
def scroll():
    elements = driver.find_elements(By.CLASS_NAME, '_1k5soqfl')
    count_el = len(elements)
    count_el2 = 0
    while count_el != count_el2:
        count_el = len(elements)
        last_element = elements[-1]
        driver.execute_script("arguments[0].scrollIntoView(true);", last_element)
        sleep(4)  # Даем время для загрузки новых элементов
        elements = driver.find_elements(By.CLASS_NAME, '_1k5soqfl')
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
review_blocks = soup.find_all(class_='_1k5soqfl')
# print(review_blocks)
# Инициализация структуры данных
keys = {'№': [], 'Имя': [], 'Рейтинг': [], 'Комментарий': []}

# Парсим данные из каждого блока
for idx, block in enumerate(review_blocks, start=1):
    keys['№'].append(idx)  # Номер отзыва

    # Имя комментатора
    name = block.find('span', {'class': '_16s5yj36'})
    keys['Имя'].append(name['title'].strip() if name else 'Ошибка')
# комментарий
try:
    # Ищем все элементы <a> с классом '_h3pmwn', где находится текст комментария
    comment_tags = block.find_all('a', {'class': '_h3pmwn'})
    
    # Если найден хотя бы один элемент, извлекаем первый комментарий
    if comment_tags:
        comment_text = comment_tags[0].text.strip()  # Убираем лишние пробелы
        keys['Комментарий'].append(comment_text)
    else:
        keys['Комментарий'].append('Ошибка')
except Exception as e:
    print(f"Ошибка при обработке текста отзыва: {e}")
    keys['Комментарий'].append('Ошибка')

# Рейтинг
try:
    # Ищем все элементы <svg> с желтыми звездами
    star_tags = block.find_all('svg', {'fill': '#ffb81c'})  # Звезды заполненные золотым цветом
    
    # Количество звезд
    rating_value = len(star_tags)  # Количество найденных звезд
    keys['Рейтинг'].append(rating_value)
except Exception as e:
    print(f"Ошибка при обработке рейтинга: {e}")
    keys['Рейтинг'].append('Ошибка')

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
print("Все отзывы успешно сохранены")
