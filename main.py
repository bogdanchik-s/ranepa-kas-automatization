import re
import traceback
import time
import json
import logging

from selenium import webdriver
from selenium.webdriver.edge.webdriver import WebDriver, Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


logging.basicConfig(
    filename='log.log',
    filemode='a',
    encoding='utf-8',
    datefmt='%d.%m.%Y %H:%M:%S',
    format='[%(asctime)s] [%(levelname)s] | %(message)s',
    level=logging.INFO
)


def wait_and_find_element(driver, locator: tuple[str, str]):
    element = WebDriverWait(driver, 500).until(EC.presence_of_element_located(locator))
    time.sleep(0.1)
    return element


def login(driver: WebDriver, username, password):
    login_field = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.ID, 'userName'))
    )
    password_field = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.ID, 'userPassword'))
    )

    login_field.send_keys(username)
    password_field.send_keys(password)

    driver.find_element(By.ID, 'okButton').click()


def write_string_to_file(*substrings):
    with open('contracts2.txt', mode='a', encoding='utf-8') as file:
        file.write('|'.join('[Пусто]' if substring is None or len(substring.strip()) == 0 else substring for substring in substrings) + '\n')


def run(driver: WebDriver):
    logging.info('Работа с договорами начата успешно.')

    actions = ActionChains(driver)

    contract_table = driver.find_element(By.ID, 'grid_form2_Список').find_element(By.CLASS_NAME, 'gridBody')
    scroll_btn = driver.find_element(By.ID, 'vertButtonScroll_form2_Список').find_element(By.CLASS_NAME, 'turnUp')

    page_num = 1
    form_num = 2

    while True:
        contract_lines = WebDriverWait(driver, 500).until(EC.presence_of_all_elements_located(locator=(By.XPATH, "//div[@id='grid_form2_Список']/div[contains(@class, 'gridBody')]/div[contains(@class, 'gridLine')]")))

        for line_idx in range(len(contract_lines)):
            contract_line = wait_and_find_element(driver, (By.XPATH, f"//div[@id='grid_form2_Список']/div[contains(@class, 'gridBody')]/div[contains(@class, 'gridLine') and position()={line_idx+1}]"))

            actions.double_click(contract_line).perform()

            page_num += 1
            form_num += 1

            # Ждем загрузки страницы
            wait_and_find_element(driver, (By.ID, f'page{page_num}'))

            close_btn = wait_and_find_element(driver, (By.ID, f'VW_page{page_num}headerTopLine_cmd_CloseButton'))

            # student_table = wait_and_find_element(driver, (By.ID, f'grid_form{form_num}_Обучающиеся')).find_element(By.CLASS_NAME, 'gridBody')
            student_table = wait_and_find_element(driver, (By.XPATH, '(//div[starts-with(@id, "form") and contains(@id, "_Обучающиеся")])[last()]')).find_element(By.CLASS_NAME, 'gridBody')
            contract_student = student_table.find_element(By.CLASS_NAME, 'gridLine').find_elements(By.CLASS_NAME, 'gridBox')[1].text.strip()
            # contract_prof = wait_and_find_element(driver, (By.ID, f'form{form_num}_ОсновнаяНоменклатура_i0')).get_property('value').strip()
            contract_prof = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ОсновнаяНоменклатура_i0")])[last()]')).get_property('value').strip()
            # contract_code = wait_and_find_element(driver, (By.ID, f'form{form_num}_Код_i0')).get_property('value').strip()
            contract_code = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Код_i0")])[last()]')).get_property('value').strip()
            # contract_datetime = wait_and_find_element(driver, (By.ID, f'form{form_num}_Дата_i0')).get_property('value').strip()
            contract_datetime = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Дата_i0")])[last()]')).get_property('value').strip()

            contract_code = driver.find_elements(By.ID, re.compile(r'form\d+_Код_i0'))[-1]
            print(contract_code)

            input('Пауза')

            # Открываем заказчика
            # driver.find_element(By.ID, f'form{form_num}_кнПросмотрКонтрагента').click()
            wait_and_find_element(driver, (By.XPATH, '(//a[starts-with(@id, "form") and contains(@id, "_кнПросмотрКонтрагента")])[last()]')).click()

            form_num += 1

            wait_and_find_element(driver, (By.ID, f'VW_page{page_num}ps0win'))

            # Открываем выпадающий список Еще
            # driver.find_element(By.ID, f'form{form_num}_allActionsФормаКоманднаяПанель').click()
            # driver.find_element(By.XPATH, '(//a[starts-with(@id, "form") and contains(@id, "_allActionsФормаКоманднаяПанель")])[last()]').click()
            wait_and_find_element(driver, (By.XPATH, '(//a[starts-with(@id, "form") and contains(@id, "_allActionsФормаКоманднаяПанель")])[last()]')).click()

            # Открываем список контрагентов
            # driver.find_element(By.ID, f'form{form_num}_popup_ФормаПоказатьВСписке').click()
            # driver.find_element(By.XPATH, '(//div[starts-with(@id, "form") and contains(@id, "_popup_ФормаПоказатьВСписке")])[last()]').click()
            wait_and_find_element(driver, (By.XPATH, '(//div[starts-with(@id, "form") and contains(@id, "_popup_ФормаПоказатьВСписке")])[last()]')).click()

            page_num += 1
            form_num += 1

            wait_and_find_element(driver, (By.ID, f'page{page_num}'))

            # Открываем карточку контрагента из списка
            driver.switch_to.active_element.send_keys(Keys.ENTER)

            form_num += 1

            # contract_customer_code = wait_and_find_element(driver, (By.ID, f'form{form_num}_Код_i0')).get_property('value').strip()
            contract_customer_code = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Код_i0")])[last()]')).get_property('value').strip()
            # contract_customer_name = wait_and_find_element(driver, (By.ID, f'form{form_num}_Наименование_i0')).get_property('value').strip()
            contract_customer_name = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Наименование_i0")])[last()]')).get_property('value').strip()
            # contract_customer_individual_name = wait_and_find_element(driver, (By.ID, f'form{form_num}_ФизЛицо_i0')).get_property('value').strip()
            contract_customer_individual_name = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ФизЛицо_i0")])[last()]')).get_property('value').strip()

            if len(contract_customer_individual_name) > 0:
                # Вставляем ФИО в графу Полное наименование
                # contract_customer_fullname = wait_and_find_element(driver, (By.ID, f'form{form_num}_НаименованиеПолное_i0'))
                contract_customer_fullname = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_НаименованиеПолное_i0")])[last()]'))
                contract_customer_fullname.clear()
                contract_customer_fullname.send_keys(contract_customer_individual_name)

                # Указываем тип физического лица
                # contract_customer_type = wait_and_find_element(driver, (By.ID, f'form{form_num}_ТипКонтрагента_i0'))
                contract_customer_type = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ТипКонтрагента_i0")])[last()]'))
                contract_customer_type.clear()
                contract_customer_type.send_keys('Физическое лицо')
                contract_customer_type.send_keys(Keys.ENTER)

                # Указываем папку, где будет контрагент
                # contract_customer_folder = wait_and_find_element(driver, (By.ID, f'form{form_num}_Родитель_i0'))
                contract_customer_folder = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Родитель_i0")])[last()]'))
                contract_customer_folder.clear()
                contract_customer_folder.send_keys('КМПО 2024')
                contract_customer_folder.send_keys(Keys.ENTER)

                # Вписываем место обучения
                # contract_customer_study_place = wait_and_find_element(driver, (By.ID, f'form{form_num}_МестоОбучения_i0'))
                contract_customer_study_place = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_МестоОбучения_i0")])[last()]'))
                contract_customer_study_place.clear()
                contract_customer_study_place.send_keys('Колледж многоуровневого профессионального образования')
                contract_customer_study_place.send_keys(Keys.ENTER)

                # Сохраняем изменения в контрагенте
                # actions.move_to_element(driver.find_element(By.ID, f'form{form_num}_ФормаЗаписать')).click().perform()
                wait_and_find_element(driver, (By.XPATH, '(//a[starts-with(@id, "form") and contains(@id, "_ФормаЗаписать")])[last()]')).click()

                # Закрываем информационное окошко
                wait_and_find_element(driver, (By.XPATH, "//div[contains(@class, 'confirm')]/span[@title='Закрыть']")).click()

                # Открываем физцило заказчика
                # driver.find_element(By.ID, f'form{form_num}_ФизЛицо_OB').click()
                # driver.find_element(By.XPATH, '(//span[starts-with(@id, "form") and contains(@id, "_ФизЛицо_OB")])[last()]').click()
                wait_and_find_element(driver, (By.XPATH, '(//span[starts-with(@id, "form") and contains(@id, "_ФизЛицо_OB")])[last()]')).click()

                form_num += 1

                wait_and_find_element(driver, (By.ID, f'VW_page{page_num}ps1win'))
                
                print(f'Страница физлица контрагента открыта [{contract_customer_name}]. PAGE_NUM: {page_num} FORM_NUM: {form_num}')

                contract_customer_individual_code = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Код_i0")])[last()]')).get_property('value').strip()
                contract_customer_group = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Родитель_i0")])[last()]')).get_property('value').strip()
                contract_customer_birthdate = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДатаРождения_i0")])[last()]')).get_property('value').strip()

                # Проверяем наличие заполненного гражданства
                contract_customer_nationality = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_Гражданство_i0")])[last()]')).get_property('value').strip()

                # Проверяем наличие паспортных данных
                contract_customer_document = wait_and_find_element(driver, (By.XPATH, '(//div[starts-with(@id, "form") and contains(@id, "_ГруппаПаспортныеДанные#title_div")])[last()]'))

                if contract_customer_document.is_displayed():
                    contract_customer_document_series = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДокументСерия_i0")])[last()]')).get_property('value')
                    contract_customer_document_number = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДокументНомер_i0")])[last()]')).get_property('value')
                    contract_customer_document_department = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДокументКемВыдан_i0")])[last()]')).get_property('value')
                    contract_customer_document_code_department = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДокументКодПодразделения_i0")])[last()]')).get_property('value')
                    contract_customer_document_date = wait_and_find_element(driver, (By.XPATH, '(//input[starts-with(@id, "form") and contains(@id, "_ДокументДатаВыдачи_i0")])[last()]')).get_property('value')
                else:
                    contract_customer_document_series = None
                    contract_customer_document_number = None
                    contract_customer_document_department = None
                    contract_customer_document_code_department = None
                    contract_customer_document_date = None

                # Закрываем карточку фтзлица заказчика
                wait_and_find_element(driver, (By.ID, f'VW_page{page_num}ps1headerTopLine_cmd_CloseButton')).click()
            else:
                contract_customer_individual_code = 'НЕТ ФИЗЛИЦА'
                contract_customer_group = 'НЕТ ФИЗЛИЦА'
                contract_customer_birthdate = 'НЕТ ФИЗЛИЦА'
                contract_customer_nationality = 'НЕТ ФИЗЛИЦА'

                contract_customer_document_series = 'НЕТ ФИЗЛИЦА'
                contract_customer_document_number = 'НЕТ ФИЗЛИЦА'
                contract_customer_document_department = 'НЕТ ФИЗЛИЦА'
                contract_customer_document_code_department = 'НЕТ ФИЗЛИЦА'
                contract_customer_document_date = 'НЕТ ФИЗЛИЦА'

            # Записываем полученные данные в файл
            write_string_to_file(
                contract_code, contract_datetime,
                contract_student,
                contract_customer_code, contract_customer_individual_code, contract_customer_group, contract_customer_name, contract_customer_birthdate, contract_customer_nationality,
                contract_customer_document_series, contract_customer_document_number, contract_customer_document_department, contract_customer_document_code_department, contract_customer_document_date
            )

            logging.info(f'Добавлен договор №{contract_code} от {contract_datetime}. Обучающийся: {contract_student}. Контрагент: {contract_customer_name} ({contract_customer_code})')

            # Закрываем карточку контрагента из списка
            wait_and_find_element(driver, (By.ID, f'VW_page{page_num}ps0headerTopLine_cmd_CloseButton')).click()

            # Закрываем список контрагентов
            wait_and_find_element(driver, (By.ID, f'VW_page{page_num}headerTopLine_cmd_CloseButton')).click()

            # Закрываем карточку заказчика
            wait_and_find_element(driver, (By.ID, f'VW_page{page_num-1}ps0headerTopLine_cmd_CloseButton')).click()

            # Закрываем договор
            close_btn.click()
        
        if 'disabled' not in scroll_btn.get_attribute('class'):
            actions.move_to_element(scroll_btn).click().perform()
        else:
            input('Был обнаружен конец таблицы. Для завершения работы нажмите Enter')
            logging.info('Был обнаружен конец таблицы.')
            break


if __name__ == '__main__':
    logging.info('Скрипт успешно запущен')

    driver = webdriver.Edge()
    driver.implicitly_wait(15)
    driver.maximize_window()

    try:
        driver.get('https://kas.ranepa.ru/kas/ru_RU/')

        login(driver, 'Захаров Богдан Сергеевич', '99IbsKTq')

        input('После открытия нужного раздела нажмите Enter')

        run(driver)
    except Exception as e:
        print('CRITICAL:\n' + traceback.format_exc())
        logging.critical(traceback.format_exc())
    finally:
        logging.info('Завершение работы...')
        driver.close()
        driver.quit()
