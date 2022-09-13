import datetime
import re
import smtplib
import telebot
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from auth import *

# создание бота
bot = telebot.TeleBot(token)

def PrintOffice24Order(printoffice24_order_url, message):
    # параметры для открытия браузера
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=3840,2160")
    chrome_options.add_argument("--no-sandbox")
    browser = webdriver.Chrome(executable_path='/Users/pavel/Documents/Install_mac/chromedriver/chromedriver',
                               options=chrome_options)

    # авторизация на printoffice24
    browser.get("https://printoffice24.com/login")

    printoffice24_login_xpath = '/html/body/div[3]/section/div/div/div/form/div[1]/div/input'
    browser.find_element(By.XPATH, printoffice24_login_xpath).send_keys(printoffice24_login)

    printoffice24_password_xpath = '/html/body/div[3]/section/div/div/div/form/div[2]/div/input'
    browser.find_element(By.XPATH, printoffice24_password_xpath).send_keys(printoffice24_password)

    printoffice24_entry_button_xpath = '/html/body/div[3]/section/div/div/div/form/div[3]/div/button'
    browser.find_element(By.XPATH, printoffice24_entry_button_xpath).click()

    print('Залогинился на PrintOffice24...')
    # bot.send_message(message.chat.id, 'Залогинился на PrintOffice24...')

    # открытие заказа
    browser.get(printoffice24_order_url)

    # поиск id элемента сделки
    deal_xpath = "//tr[@class='deal-element-row']"
    deal_id = []
    for i in browser.find_elements(By.XPATH, deal_xpath):
        deal_id.append(i.get_attribute('id'))  # deal_elem_3640977
    order = []

    # открытие заказа
    browser.get(printoffice24_order_url)
    time.sleep(0.1)

    # в каждой строчке заказа собираем информацию, если стоит статус Meduzz
    for i in deal_id:

        if 'Статус: Meduzz' in browser.find_element(By.XPATH, f"//tr[@id='{i}']").text:
            print('Нашел заказ со статусом Meduzz...')
            # bot.send_message(message.chat.id, 'Нашел заказ со статусом Meduzz...')
            temp = []

            # получение названия продукции (полное) ORDER
            printoffice24_order_description_xpath = f"//tr[@id='{i}']/td/span[@class='deal-element-name']"
            printoffice24_order_description = browser.find_element(By.XPATH, printoffice24_order_description_xpath).text
            temp.append(printoffice24_order_description)
            print(f'Название продукции полное: {printoffice24_order_description}')
            # bot.send_message(message.chat.id, f'Название продукции полное: {printoffice24_order_description}')

            # получение названия продукции (краткое) PRODUCT
            try:
                printoffice24_product_description = f'{printoffice24_order_description.split()[0]} {printoffice24_order_description.split()[1]} {printoffice24_order_description.split()[2]} {printoffice24_order_description.split()[3]}'
            except:
                printoffice24_product_description = f'{printoffice24_order_description.split()[0]} {printoffice24_order_description.split()[1]}'
            temp.append(printoffice24_product_description)
            print(f'Название продукции краткое: {printoffice24_product_description}')
            # bot.send_message(message.chat.id, f'Название продукции краткое: {printoffice24_product_description}')

            # получение количества листов А3 (между двумя звездочками * количество листов *)
            try:
                printoffice24_number_of_A3_sheets_xpath = f"//tr[@id='{i}']/td/p[@class='deal-elm-description']"
                printoffice24_number_of_A3_sheets = re.search(r'\*(.*?)\*', browser.find_element(By.XPATH,
                                                                                                 printoffice24_number_of_A3_sheets_xpath).text).group(
                    1)
                temp.append(printoffice24_number_of_A3_sheets)
            except:
                printoffice24_number_of_A3_sheets = '0'
                temp.append(printoffice24_number_of_A3_sheets)
                bot.send_message(message.chat.id, "Ты забыл поставить *количество листов* !!!\nУстановлено по умолчанию 0.")
            print(f'Количество листов: {printoffice24_number_of_A3_sheets}')
            # bot.send_message(message.chat.id, f'Количество листов: {printoffice24_number_of_A3_sheets}')

            # получение количества изделий
            printoffice24_quantity_xpath = f"//tr[@id='{i}']/td/span[@class='autonumber-positive']"
            printoffice24_quantity = browser.find_element(By.XPATH, printoffice24_quantity_xpath).text.replace("\u2009",
                                                                                                               "")
            temp.append(printoffice24_quantity)
            print(f'Количество изделий: {printoffice24_quantity}')
            # bot.send_message(message.chat.id, f'Количество изделий: {printoffice24_quantity}')

            # нажатие на кнопку eye для открытия цен себестоимости
            printoffice24_eye_xpath = "//button[@class='btn btn-xs btn-default p-0-6 show-hidden-values']"
            browser.find_element(By.XPATH, printoffice24_eye_xpath).click()
            time.sleep(0.1)

            # получение цены
            printoffice24_price_xpath = f"//tr[@id='{i}']/td/span[@class='autonumber-positive hidden-values']"
            printoffice24_price = browser.find_element(By.XPATH, printoffice24_price_xpath).text.replace("\u2009", "")
            temp.append(printoffice24_price)
            print(f'Цена, которая будет в Meduzz: {printoffice24_price}')
            bot.send_message(message.chat.id, f'Проверь!\nЦена, которая будет в Meduzz: {printoffice24_price}')

            # получение названия компании
            printoffice24_company_name_xpath = f"//div[@class='about-info-p']/span"
            printoffice24_company_name = browser.find_element(By.XPATH, printoffice24_company_name_xpath).text
            temp.append(printoffice24_company_name)
            print(f'Заказчик: {printoffice24_company_name}')
            # bot.send_message(message.chat.id, f'Заказчик: {printoffice24_company_name}')

            # получение номера заказа
            printoffice24_order_number_xpath = "//div[@id='deal_info']"
            printoffice24_order_number = browser.find_element(By.XPATH, printoffice24_order_number_xpath).get_attribute(
                'data-num')
            print(f'Номер заказа в PrintOffice24: {printoffice24_order_number}')
            # bot.send_message(message.chat.id, f'Номер заказа в PrintOffice24: {printoffice24_order_number}')

            # получение списка загруженных макетов
            printoffice24_layout_button_xpath = f"//tr[@id='{i}']/td[4]/a[@class='btn btn-xs btn-primary waves-effect waves-light deal_layouts_cnt']"
            browser.find_element(By.XPATH, printoffice24_layout_button_xpath).click()
            time.sleep(1)

            files_name = []
            printoffice24_layout_xpath = "//a[@target='_blank']"
            selemium_files_name = browser.find_elements(By.XPATH, printoffice24_layout_xpath)
            for item in selemium_files_name:
                if item.text != '':
                    files_name.append(item.text)
            webdriver.ActionChains(browser).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
            print(f'Имена файлов для печати: {files_name}')
            # bot.send_message(message.chat.id, f'Имена файлов для печати: {files_name}')

            # установка статуса Meduzz на статус "Статус не выбран"
            printoffice24_status_xpath = f"//tr[@id='{i}']//a[@class='btn btn-xs btn-primary on-default edit-row modal_element_status p_tooltip']/i"
            browser.find_element(By.XPATH, printoffice24_status_xpath).click()
            time.sleep(1)

            printoffice24_status_meduzz_xpath = "//button[@title='Meduzz']"
            browser.find_element(By.XPATH, printoffice24_status_meduzz_xpath).click()
            time.sleep(1)

            printoffice24_status_no_select_xpath = "/html/body/div[1]/div[1]/div[6]/div/div/div/div[2]/div/div/form/div[1]/div/div/div/div/ul/li[1]/a/span[2]"
            browser.find_element(By.XPATH, printoffice24_status_no_select_xpath).click()
            time.sleep(1)

            printoffice24_status_meduzz_xpath = "//button[@id='deal_element_status_date_save']"
            browser.find_element(By.XPATH, printoffice24_status_meduzz_xpath).click()
            time.sleep(1)

            # Отправка ссылки
            links = []
            try:
                for name in files_name:
                    tempLink = y.get_download_link(f"/printoffice24/layouts/{printoffice24_order_number}/{name}")
                    links.append(tempLink)
                    # bot.send_message(message.chat.id, f'Макет: {tempLink}')
                    # tempLink = ""
                for link in links:
                    files_name.append(link)
                if len(links) != 0:
                    temp.append(files_name)
                else:
                    noLink = "Файлы не загружены в CRM, возможно я отправил тебе их другим письмом или другим образом.\nСвяжись со мной для уточнения деталей\n\nhttps://api.whatsapp.com/send?phone=79263000085"
                    files_name.append(noLink)
                    temp.append(files_name)
            except:
                print("Ошибка добавления ссылки")
                bot.send_message(message.chat.id, "Ошибка добавления ссылки!!!")

            # Добавление к общему заказу 1 позиции
            order.append(temp)

            # открытие заказа
            browser.get(printoffice24_order_url)
            time.sleep(0.5)

        else:
            print("В нескольких позициях не обнаружено статуса Meduzz.")
            # bot.send_message(message.chat.id, 'В нескольких позициях не обнаружено статуса Meduzz.')

    if order == []:
        bot.send_message(message.chat.id, 'Ошибка: для отправки заказа установи статус "Meduzz" хотя бы для одной из позицй')

    # часть функции относящаяся к Meduzz:
    for i in order:  # i = ['Диплом А4, 4+0, бумага DNS 300 гр./м2 / Цифровая печать', 'Диплом А4, 4+0, бумага',      '5',          '10',       '168',        'YOJI']
        #      --------------Полное описание[0]-------------------------  --------краткое[1]-------  --листов[2]-- --кол-во[3]-- --цена[4]-- --компания[5]--

        # открыть браузер
        browser.get('https://printhq.ru/crm/login.php')

        # заполнить логин
        meduzz_login_xpath = '/html/body/div[3]/div/div/div/div/form/table/tbody/tr[1]/td/div/input'
        browser.find_element(By.XPATH, meduzz_login_xpath).send_keys(meduzz_login)

        # заполнить пароль
        meduzz_password_xpath = '/html/body/div[3]/div/div/div/div/form/table/tbody/tr[2]/td/div/input'
        browser.find_element(By.XPATH, meduzz_password_xpath).send_keys(meduzz_password)

        # нажать на вход
        meduzz_entry_button_xpath = '/html/body/div[3]/div/div/div/div/form/table/tbody/tr[3]/td/input'
        browser.find_element(By.XPATH, meduzz_entry_button_xpath).click()

        # перейти на добавление нового заказа в meduzz
        browser.get('https://printhq.ru/crm/index.php?act=add')

        # заполнение сегодняшнего числа
        date_today = datetime.datetime.today().strftime('%d.%m.%Y')
        date_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[1]/td/table/tbody/tr/td[3]/input'
        browser.find_element(By.XPATH, date_xpath).send_keys(date_today)

        # заполнение заказчика
        meduzz_customer_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[2]/td/table[1]/tbody/tr/td[2]/select'
        browser.find_element(By.XPATH, meduzz_customer_xpath).send_keys('Зиборов Павел')

        # заполнение типа устройства
        meduzz_printer_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[3]/td/table[1]/tbody/tr/td[1]/select'
        browser.find_element(By.XPATH, meduzz_printer_xpath).send_keys('Не выбрано')

        # заполнение типа продукции
        meduzz_products_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[4]/td/table[1]/tbody/tr/td[1]/select'
        browser.find_element(By.XPATH, meduzz_products_xpath).send_keys('В доп. поле')

        # заполнение описание заказа (полное)
        meduzz_order_description_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[4]/td/table[3]/tbody/tr/td/textarea'
        meduzz_order_description = i[0]
        browser.find_element(By.XPATH, meduzz_order_description_xpath).send_keys(meduzz_order_description)

        # заполнение описания продукции (краткое)
        meduzz_product_description_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[4]/td/table[1]/tbody/tr/td[2]/input'
        try:
            meduzz_product_description = f'{i[5].split()[0]} | {i[1]}'
            browser.find_element(By.XPATH, meduzz_product_description_xpath).send_keys(meduzz_product_description)
        except:
            meduzz_product_description = f'{i[5]} | {i[1]}'
            browser.find_element(By.XPATH, meduzz_product_description_xpath).send_keys(meduzz_product_description)

        # заполнение количества листов
        meduzz_number_of_A3_sheets_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[4]/td/table[1]/tbody/tr/td[3]/input'
        browser.find_element(By.XPATH, meduzz_number_of_A3_sheets_xpath).send_keys(i[2])

        # заполнение количества изделий
        meduzz_quantity_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[4]/td/table[2]/tbody/tr/td[2]/input'
        browser.find_element(By.XPATH, meduzz_quantity_xpath).send_keys(i[3])

        # заполнение тела письма
        meduzz_dop_description_xpath = "//textarea[@name='add_description_to_the_order']"
        # browser.find_element(By.XPATH, meduzz_dop_description_xpath).send_keys(meduzz_dop_description)

        # заполнение цены в медузе
        meduzz_price_xpath = '/html/body/div[3]/div/div[2]/form/table/tbody/tr[16]/td/table/tbody/tr/td[1]/input'
        meduzz_price = i[4]
        browser.find_element(By.XPATH, meduzz_price_xpath).send_keys(meduzz_price)

        # нажатие на кнопку сохранения
        meduzz_save_button_xpath = "//input[@type='submit']"
        browser.find_element(By.XPATH, meduzz_save_button_xpath).click()
        print("Создал заказ в Meduzz...")

        # открыть браузер
        browser.get('https://printhq.ru/crm/index.php')
        time.sleep(0.5)

        # узнать номер заказа
        order_number_xpath = "//td[@class='sorting_1']"
        meduzz_order_number = browser.find_element(By.XPATH, order_number_xpath).text.split("/")[0]
        print(f'Номер заказа в Meduzz: {meduzz_order_number}')
        # bot.send_message(message.chat.id, f'Номер заказа в Meduzz: {meduzz_order_number}')

        # отправка письма Андрею
        meduzz_dop_description = "Список файлов для заказа и ссылка ниже:\n\n"
        for j in i[6]:
            meduzz_dop_description = meduzz_dop_description + j + "\n\n"
        SUBJECT = f"{meduzz_order_number} | {meduzz_order_description}"
        text = f"{meduzz_dop_description}"
        SendEmail(SUBJECT, TO, FROM, text)
        print(f'Отправлено на почту печатнику...')
        # bot.send_message(message.chat.id, f'Отправлено на почту печатнику...')

        # отправка копии письма Себе
        SendEmail(SUBJECT, "pavel@heavenprint.ru", FROM, text)
        print(f'Отправлена копия на pavel@heavenprint.ru...')
        # bot.send_message(message.chat.id, f'Отправлена копия на pavel@heavenprint.ru...')

        # сообщение для отправки в telegram
        telegram_send_files = "\n\n"
        for j in i[6]:
            telegram_send_files = telegram_send_files + j + "\n\n"
        bot.send_message(message.chat.id, f'\n\nЗаказ отправлен в печать.\n'
                                          f'\nНомер заказа PrintOffice: {printoffice24_order_number}'
                                          f'\nНомер заказа Meduzz: {meduzz_order_number}'
                                          f'\nЗаказчик: {printoffice24_company_name}'
                                          f'\nОписание: {printoffice24_order_description}'
                                          f'\nЦена в Meduzz: {meduzz_price}')
        bot.send_message(message.chat.id, f'\nСсылка на макет(ы): {telegram_send_files}')

    browser.close()


def SendEmail(SUBJECT, TO, FROM, text):
    HOST = "smtp.mail.ru"
    login = 'pavel@heavenprint.ru'
    password = 'CQZJNCLbypjFRhbepFhM'

    BODY = "\r\n".join((
        "From: %s" % FROM,
        "To: %s" % TO,
        "Subject: %s" % SUBJECT,
        "",
        text
    ))

    server = smtplib.SMTP_SSL(HOST)
    server.login(login, password)
    server.sendmail(FROM, [TO], BODY.encode('utf-8'))
    server.quit()


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет!")


@bot.message_handler(content_types=['text'])
def send_order(message):
    if 'https://printoffice24.com/editDeal/' in message.text:
        bot.send_message(message.chat.id, "Начинаю отправку заказа...")
        PrintOffice24Order(message.text, message)
    else:
        bot.send_message(message.chat.id, "Ошибка: не верный формат ссылки")



print('Start server')
bot.infinity_polling()
