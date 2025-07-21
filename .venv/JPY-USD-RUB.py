from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import numbers
import xml.etree.ElementTree as ET
import time
from selenium.webdriver.chrome.options import Options
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# Настройки для Chrome
download_dir = os.path.abspath("C:/ReportMOEX")
os.makedirs(download_dir, exist_ok=True)


chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,  # Основная папка
    "download.prompt_for_download": False,       # Отключить диалог "Сохранить как"
    "download.directory_upgrade": True,          # Разрешить загрузку в указанную папку
    "safebrowsing.enabled": True                # Безопасный режим (опционально)
})

driver = webdriver.Chrome(options=chrome_options)

# 1. Открыть сайт MOEX
  # Запускаем браузер Chrome
driver.maximize_window()
driver.get("https://www.moex.com")

# 2. Навести на "Срочный рынок" и перейти в "Индикативные курсы"
try:
    # Ждём, пока загрузится страница и находим элемент "Срочный рынок"
    futures_market = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Срочный рынок"))
    )

    time.sleep(1)

    # Наводим курсор на элемент
    ActionChains(driver).move_to_element(futures_market).perform()
    time.sleep(1)  # Даем время для появления подменю

    # Кликаем на "Индикативные курсы" в выпадающем меню
    indicative_rates = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Индикативные курсы"))
    )
    indicative_rates.click()
    time.sleep(1)
    # Соглашаемся
    accept_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Согласен"))
    )
    accept_button.click()
    time.sleep(1)
    # Открываем список
    accept_button = driver.find_element(By.CLASS_NAME, "ui-select__placeholder")
    accept_button.click()
    time.sleep(1)
    # Выбираем USD/RUB - Доллар США к российскому рублю
    accept_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "USD/RUB - Доллар США к российскому рублю"))
    )
    accept_button.click()
    time.sleep(1)


    # Выбираем период с и по

    # Открываем календарь(первый)
    label = driver.find_element(By.CSS_SELECTOR, "label[for='fromDate']")
    label.click()
    time.sleep(1)

    # Открываем список
    month_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "ui-calendar__controls"))
    )
    month_dropdown.click()
    time.sleep(1)

    # # Выбираем в списке июнь
    month_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-select-option__content"))
    )

    for month in month_options:
        if month.text == "06 - Июнь":
            month.click()
            break
    time.sleep(1)

    #выбираем первое число
    day_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-calendar__cell"))
    )

    for day in day_options:
        if day.text == "1":
            day.click()
            break

    time.sleep(2)

    # Открываем календарь(второй)
    label = driver.find_element(By.CSS_SELECTOR, "label[for='tillDate']")
    label.click()
    time.sleep(1)

    # Открываем список

    all_calendar_controls = driver.find_elements(By.CSS_SELECTOR, ".ui-calendar__controls")
    second_calendar_control = all_calendar_controls[1]
    second_calendar_control.click()
    time.sleep(1)


    # # Выбираем в списке июнь
    month_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-select-option__content"))
    )

    for month in month_options:
        if month.text == "06 - Июнь":
            month.click()
            break

    time.sleep(1)


    #выбираем 30 число

    day_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-calendar__cell"))
    )

    for day in day_options:
        if day.text == "30":
            day.click()
            break

    time.sleep(1)
    #Показать

    all_show = driver.find_elements(By.CSS_SELECTOR, ".ui-button__label")
    show_me = all_show[1]
    show_me.click()
    time.sleep(1)

    time.sleep(1)


    #получаем данные
    # 5. Находим и кликаем на кнопку "XML" (она может быть в выпадающем меню)
    xml_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'XML') or contains(@class, 'export-xml')]"))
    )
    xml_button.click()
    time.sleep(10)

    driver.quit()
except Exception as e:
    print("Ошибка при переходе:", e)
    driver.quit()



xml_path = r'C:\ReportMOEX\currencyRate-USD_RUB-20250601-20250630.xml'
if not os.path.exists(xml_path):
    raise FileNotFoundError(f"Файл не найден: {xml_path}")
# Загрузка XML-файла
tree = ET.parse(r'C:\ReportMOEX\currencyRate-USD_RUB-20250601-20250630.xml')
root = tree.getroot()

# Создание Excel-файла
wb = Workbook()
ws = wb.active
ws.title = "USD_RUB_Rates"

# Заголовки столбцов
ws['A1'] = 'Дата USD/RUB'
ws['B1'] = 'Курс USD/RUB'
ws['C1'] = 'Время USD/RUB'

# Обработка данных из XML
row_num = 2  # Начинаем с второй строки (первая - заголовки)
for row in root.findall('.//rows/row'):
    tradetime = row.get('tradetime')
    if tradetime == "18:49:00":  # Фильтр по времени
        ws.cell(row=row_num, column=1, value=row.get('tradedate'))
        ws.cell(row=row_num, column=2, value=float(row.get('rate')))
        ws.cell(row=row_num, column=3, value=tradetime)
        row_num += 1

# Автонастройка ширины столбцов
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

# Сохранение Excel-файла
wb.save(r'C:\ReportMOEX\USD_RUB_Rates.xlsx')
print("Данные успешно экспортированы в USD_RUB_Rates.xlsx")



driver = webdriver.Chrome(options=chrome_options)

# 1. Открыть сайт MOEX
  # Запускаем браузер Chrome
driver.maximize_window()
driver.get("https://www.moex.com")

# 2. Навести на "Срочный рынок" и перейти в "Индикативные курсы"
try:
    # Ждём, пока загрузится страница и находим элемент "Срочный рынок"
    futures_market = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Срочный рынок"))
    )

    time.sleep(1)

    # Наводим курсор на элемент (вместо клика)
    ActionChains(driver).move_to_element(futures_market).perform()
    time.sleep(1)  # Даем время для появления подменю

    # Кликаем на "Индикативные курсы" в выпадающем меню
    indicative_rates = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Индикативные курсы"))
    )
    indicative_rates.click()
    time.sleep(1)
    # Соглашаемся
    accept_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Согласен"))
    )
    accept_button.click()
    time.sleep(1)
    # Открываем список
    accept_button = driver.find_element(By.CLASS_NAME, "ui-select__placeholder")
    accept_button.click()
    time.sleep(1)
    # Выбираем USD/RUB - Доллар США к российскому рублю
    accept_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "JPY/RUB - Японская йена к российскому рублю"))
    )
    accept_button.click()
    time.sleep(1)



    # Выбираем период с и по

    # Открываем календарь(первый)
    label = driver.find_element(By.CSS_SELECTOR, "label[for='fromDate']")
    label.click()
    time.sleep(1)

    # Открываем список
    month_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "ui-calendar__controls"))
    )
    month_dropdown.click()
    time.sleep(1)

    # # Выбираем в списке июнь
    month_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-select-option__content"))
    )

    for month in month_options:
        if month.text == "06 - Июнь":
            month.click()
            break
    time.sleep(1)

    #выбираем первое число

    day_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-calendar__cell"))
    )

    for day in day_options:
        if day.text == "1":
            day.click()
            break

    time.sleep(2)

    # Открываем календарь(второй)
    label = driver.find_element(By.CSS_SELECTOR, "label[for='tillDate']")
    label.click()
    time.sleep(1)

    # Открываем список

    all_calendar_controls = driver.find_elements(By.CSS_SELECTOR, ".ui-calendar__controls")
    second_calendar_control = all_calendar_controls[1]
    second_calendar_control.click()
    time.sleep(1)


    # Выбираем в списке июнь
    month_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-select-option__content"))
    )

    for month in month_options:
        if month.text == "06 - Июнь":
            month.click()
            break

    time.sleep(1)


    #выбираем 30 число

    day_options = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ui-calendar__cell"))
    )

    for day in day_options:
        if day.text == "30":
            day.click()
            break

    time.sleep(1)
    #Показать

    all_show = driver.find_elements(By.CSS_SELECTOR, ".ui-button__label")
    show_me = all_show[1]
    show_me.click()
    time.sleep(1)

    time.sleep(1)


    #получаем данные
    # 5. Находим и кликаем на кнопку "XML" (она может быть в выпадающем меню)
    xml_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'XML') or contains(@class, 'export-xml')]"))
    )
    xml_button.click()
    time.sleep(5)



except Exception as e:
    print("Ошибка при переходе:", e)
    driver.quit()

# Загрузка XML-файла(выбираем путь)

xml_path = r'C:\ReportMOEX\currencyRate-JPY_RUB-20250601-20250630.xml'
if not os.path.exists(xml_path):
    raise FileNotFoundError(f"Файл не найден: {xml_path}")
# Загрузка XML-файла
tree = ET.parse(r'C:\ReportMOEX\currencyRate-JPY_RUB-20250601-20250630.xml')
root = tree.getroot()
excel_path = r'C:\ReportMOEX\USD_RUB_Rates.xlsx'



# Создание Excel-файла
wb = load_workbook(excel_path)
ws = wb.active


# Заголовки столбцов
ws['D1'] = 'Дата JPY/RUB'
ws['E1'] = 'Курс JPY/RUB'
ws['F1'] = 'Время JPY/RUB'
ws['G1'] = 'Результат'

# Обработка данных из XML
row_num = 2  # Начинаем с второй строки (первая - заголовки)
for row in root.findall('.//rows/row'):
    tradetime = row.get('tradetime')
    if tradetime == "18:49:00":  # Фильтр по времени
        ws.cell(row=row_num, column=4, value=row.get('tradedate'))
        ws.cell(row=row_num, column=5, value=float(row.get('rate')))
        ws.cell(row=row_num, column=6, value=tradetime)
        row_num += 1

for row in range(2, ws.max_row + 1):  # Начинаем со 2-й строки
    value_b = ws.cell(row=row, column=2).value  # Значение из столбца B
    value_e = ws.cell(row=row, column=5).value  # Значение из столбца E

    if value_b is not None and value_e is not None and value_e != 0:
        result = value_b / value_e
        cell = ws.cell(row=row, column=7, value=result)  # Запись в столбец G (7)
        cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
# Автонастройка ширины столбцов
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width
# Числовой ли формат?
for row in range(2, ws.max_row + 1):
    cell = ws.cell(row=row, column=7)  # Столбец G
    if isinstance(cell.value, (int, float)):
        print(f"Ячейка G{row}: числовой формат ({cell.value})")
    else:
        print(f"Ячейка G{row}: НЕ числовой формат ({cell.value})")

# Сохранение Excel-файла
wb.save(r'C:\ReportMOEX\USD_JPY_RUB_Rates.xlsx')
print(r"Данные успешно экспортированы в C:\ReportMOEX\USD_JPY_RUB_Rates.xlsx")


# Делаем функцию для отправки писем
def send_excel_with_row_count(sender_email, sender_password, recipient_email, file_path):
    # Открываем Excel-файл и считаем строки без заголовков
    workbook = openpyxl.load_workbook(r'C:\ReportMOEX\USD_JPY_RUB_Rates.xlsx')
    sheet = workbook.active
    row_count = sheet.max_row - 1






    # Делаем склонение строк
    if row_count % 10 == 1 and row_count % 100 != 11:
        row_text = f"{row_count} строка"
    elif 2 <= row_count % 10 <= 4 and not (12 <= row_count % 100 <= 14):
        row_text = f"{row_count} строки"
    else:
        row_text = f"{row_count} строк"

    # Формируем письмо
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Отправка Excel-файла ({row_text})"

    body = f"""
    Добрый вечер!

    Отправляю Excel-файл с данными. В файле {row_text}.

    С уважением,
    Петров Владислав
    """
    msg.attach(MIMEText(body, 'plain',  'utf-8'))

    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename=USD_JPY_RUB_Rates',
        )
        msg.attach(part)

    try:
        server = smtplib.SMTP_SSL('smtp.yandex.ru', 465)
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print("Письмо успешно отправлено!")
    except Exception as e:
        print(f"Ошибка отправки: {e}")


# Запускаем функцию, указываем данные
send_excel_with_row_count(
    sender_email = 'vrad2019@yandex.ru',
    sender_password = os.getenv("YANDEX_PASSWORD"),
    recipient_email = 'vrad2019@yandex.ru',
    file_path = r'C:\ReportMOEX\USD_JPY_RUB_Rates.xlsx'
)

