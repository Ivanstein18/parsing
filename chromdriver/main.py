from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import xlrd
from docxtpl import DocxTemplate



url = 'https://dnevnik.ru/teachers'

options = Options()
options.add_argument('user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36')
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--headless")



chrome_prefs = {"download.default_directory": os.path.join('.', 'download')}
options.experimental_options["prefs"] = chrome_prefs


driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options= options
    )

try:
    driver.get(url= url)
    
    email_input = driver.find_element(by= By.NAME, value= 'login')
    email_input.clear()
    email_input.send_keys('******')

    password_input = driver.find_element(by= By.NAME, value= 'password')
    password_input.clear()

    with open(os.path.join('.', 'password.txt'), 'r', encoding= 'cp1251') as f:
        password = f.readline()
        password = password.split('=')[1].strip()
    if password == '':
        print('='*50)
        password = input("Введите пароль: ").strip()

    password_input.send_keys(password)

    password_input.send_keys(Keys.ENTER)

    journal2 = driver.find_elements(by= By.CLASS_NAME, value= 'header-submenu__link')[4].click()

    clas_list = driver.find_element(by= By.CLASS_NAME, value= 'classes').find_elements(by= By.TAG_NAME, value= 'a')
    week_list = driver.find_elements(by= By.XPATH, value= "//td[@class='ui-datepicker-week-col']")

    print('-'*50)
    print(f'Количество недель в месяце = {len(week_list)}')
    print('-'*50)
    print(f'Выбирите неделю от 1 до {len(week_list)}')
    print('-'*50)
    number_week = int(input())

    print("Загрузка начата!")
    print("0%")
    proc = 0

    for clas, namber in zip(clas_list[0:28], range(0, 28)):  #[0:28]

            clas.click()
            week_list = driver.find_elements(by= By.XPATH, value= "//td[@class='ui-datepicker-week-col']")
            week_list[number_week-1].click()
            
            desriptor = driver.window_handles
            driver.switch_to.window(desriptor[1])
            xls = driver.find_element(by= By.XPATH, value= "//a[@id='weekExportButton']").click()
            
            driver.close()

            desriptor = driver.window_handles
            driver.switch_to.window(desriptor[0])

            time.sleep(1)
            os.rename(os.path.join('.', 'download', 'journal.xls'), os.path.join('.', 'download', f'journal{namber}.xls'))

            proc = proc + 3.5
            print(f"{round(proc)}%")

    print("100%")
    print("Загрузка завершена!")

    time.sleep(1)


except Exception as ex:
    print(ex)

finally:
    driver.close()
    driver.quit()


#--------------------------------------------------------------------------------------------------------------------------


doc = DocxTemplate(os.path.join('.', 'tabl.docx'))
final_content = {}

print("-"*50)

for journal_number in range(0, 28):  #(0, 28)

    workbook = xlrd.open_workbook(os.path.join('.', 'download', f'journal{journal_number}.xls'))
    sheet = workbook.sheet_by_index(0)


    content = {}
    for i, col in zip(sheet.row_slice(5), range(0, 100)):
        if i.value != '':
            if journal_number == 0:
                content[f"day{(i.value).split('/')[1].strip().split('.')[0]}"] = (i.value).split('/')[1].strip().split('.')[0]
            count = 0
            for j in sheet.col_slice(col):
                if str(j.value) == 'п' or str(j.value) == 'б':
                    count = count+1
            content[f"count{journal_number}_{(i.value).split('/')[1].strip().split('.')[0]}"] = f'{int(sheet.col_slice(0)[-1].value) - count}'
            count = 0
    final_content.update(content)
    clas_name = (sheet.cell(0, 0).value).split(':')[1]
    final_content[f"clas_name_journal{journal_number}"] = clas_name

    print(f"Класс {clas_name} OK!")
    print("-"*50)

total_count = {}
for i in final_content.keys():
    if 'count' in i:
        day = i.split('_')[1]
        try:
            total_count[f"total{day}"] = int(total_count[f"total{day}"]) + int(final_content[i])
        except KeyError:
            total_count[f"total{day}"] = int(final_content[i])


final_content.update(total_count)


doc.render(final_content)
doc.save(os.path.join('.', 'tabl_final.docx'))


file_in_download = (os.listdir(path=os.path.join('.', 'download')))

for i in file_in_download:
    os.remove(os.path.join('.', 'download', f'{i}'))
