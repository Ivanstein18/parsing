from selenium import webdriver
import time
import pickle
from selenium.webdriver.common.keys import Keys



url = 'https://dnevnik.ru/teachers'

options = webdriver.ChromeOptions()
options.add_argument('user-agent= Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36')

driver = webdriver.Chrome(
    executable_path= '/home/ivanstein/Documents/Python_progect/parsing/chromdriver/chromedriver',
    options= options
    )

try:
    driver.get(url= url)
    time.sleep(1)

    # pickle.dump(driver.get_cookies(), open('cookies', 'wb'))

    # for cookie in pickle.load(open('cookies', 'rb')):
    #     driver.add_cookie(cookie)

    # time.sleep(3)
    # driver.refresh()
    # time.sleep(5)

    email_input = driver.find_element_by_name('login')
    email_input.clear()
    email_input.send_keys('m.patrikeeva@yandex.ru')

    password_input = driver.find_element_by_name('password')
    password_input.clear()
    password_input.send_keys('')

    password_input.send_keys(Keys.ENTER)

    journal2 = driver.find_elements_by_class_name('header-submenu__link')[4].click()
   
    time.sleep(5)



except Exception as ex:
    print(ex)

finally:
    driver.close()
    driver.quit()