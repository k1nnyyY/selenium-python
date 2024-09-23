import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Данные для логина
login_url = 'https://online.moysklad.ru'
purchase_order_url = 'https://online.moysklad.ru/app/#purchaseorder/edit?new'
login = 'admin@muratmynzhassar'
password = 'Osman2024'

# Чтение данных из Excel файла
file_path = 'Для Вадима.xlsx'
df = pd.read_excel(file_path, header=None)  # Загрузка без заголовков

# Настройка Selenium
driver = webdriver.Chrome()  # Убедитесь, что у вас установлен ChromeDriver
driver.get(login_url)  # Открываем страницу логина

# Ожидание, пока инпут логина станет доступен
wait = WebDriverWait(driver, 20)  # Увеличено время ожидания до 20 секунд
login_input = wait.until(EC.presence_of_element_located(
    (By.CSS_SELECTOR, 'input#lable-login')))
password_input = driver.find_element(By.CSS_SELECTOR, 'input#lable-password')

# Вводим логин и пароль
login_input.send_keys(login)
password_input.send_keys(password)

# Поиск кнопки "Войти" и клик
login_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
login_button.click()

# Ожидаем изменения URL после успешного логина
wait.until(EC.url_contains('app'))

# Переходим на страницу создания заказа
driver.get(purchase_order_url)

# Ожидание загрузки страницы с инпутом для кода
input_selector = 'input[data-test-id="consignment-selector-input"]'
input_element = wait.until(EC.presence_of_element_located(
    (By.CSS_SELECTOR, input_selector)))

# Обработка данных из Excel
for index, row in df.iterrows():
    code = row[1]  # Столбец B (код)
    count = row[2]  # Столбец C (количество)

    for _ in range(count):
        # Поиск инпута
        input_element = driver.find_element(By.CSS_SELECTOR, input_selector)
        # Вводим код
        input_element.send_keys(str(code))
        # Эмулируем нажатие Enter
        input_element.send_keys(Keys.ENTER)
        # Пауза между отправками (при необходимости)
        time.sleep(1)

# Закрытие браузера
driver.quit()
