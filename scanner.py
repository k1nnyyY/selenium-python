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
wait = WebDriverWait(driver, 20)
login_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input#lable-login')))
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
input_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, input_selector)))

# Обработка данных из Excel
for index, row in df.iterrows():
    code = row[1]  # Столбец B (код)
    count = row[2]  # Столбец C (количество)

    for _ in range(count):
        # Ожидание инпута перед каждой итерацией
        input_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, input_selector)))
        
        # Очистка поля ввода перед вводом
        input_element.clear()
        
        # Вводим код
        input_element.send_keys(str(code))
        
        # Проверка, что код действительно был введен
        if input_element.get_attribute('value') != str(code):
            print(f"Ошибка: Код {code} не был введен корректно.")
            continue
        
        # Эмулируем нажатие Enter
        input_element.send_keys(Keys.ENTER)
        
        # Пауза между отправками (при необходимости)
        time.sleep(1)

# Ввод "uniroba" в поле контрагента
contractor_input_selector = 'div[data-test-id="SystemFields.sourceAgent"] input[data-test-id="selector-input"]'
contractor_input_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, contractor_input_selector)))

# Очистка инпута перед вводом текста
contractor_input_element.clear()

# Ввод "uniroba"
contractor_input_element.send_keys('uniroba')

# Эмуляция нажатия Enter
contractor_input_element.send_keys(Keys.ENTER)

# Ожидание кнопки "Сохранить" и нажатие на неё 2-3 раза с задержкой с использованием JavaScript
save_button_selector = 'button[data-test-id="editor-toolbar-save-button"]'
for _ in range(3):  # Нажимаем на кнопку 3 раза
    save_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, save_button_selector)))
    
    # Проверяем, активна ли кнопка
    if save_button.is_enabled():
        # Используем JavaScript для нажатия на кнопку
        driver.execute_script("arguments[0].click();", save_button)
        print("Нажата кнопка 'Сохранить' с помощью JavaScript")
        time.sleep(1)  # Задержка в 1 секунду между нажатиями
    else:
        print("Кнопка 'Сохранить' не активна.")

# Ожидание сообщения об успешном сохранении
success_message_selector = 'div.gwt-Label'  # Селектор для элемента с текстом "Заказ сохранён"
success_message = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, success_message_selector)))

# Проверка текста сообщения
if "Заказ сохранён" in success_message.text:
    print("Сохранение выполнено успешно!")
else:
    print("Ошибка сохранения!")

# Закрытие браузера
driver.quit()
