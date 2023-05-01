import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pywinauto.application import Application
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl import load_workbook

if getattr(sys, 'frozen', False):
    root_path = Path(sys.executable).parent
elif __file__:
    root_path = Path(__file__).parent


BrowserHide = False
sleep_time = 1
request_time_out_seconds = 20
selenium_connect_method = 2
your_liara_env_token = 'irOrJBbi64iv5hRAwwM'
liara_chrome_app_url = 'https://zargham.iran.liara.run/webdriver'
api_key = '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'
debug = True
chromedriver_path = ''
URL = "https://fastgpt.app//"

# chromedriver_path = "C:/Users/zargham/.cache/selenium/chromedriver/win32/109.0.5414.74/chromedriver.exe"

global BrowserTitleToFind
BrowserTitleToFind = 'Zargham09355124619'


OutPut_Filename = root_path/'results.xlsx'
if OutPut_Filename.exists():
    workbook = load_workbook(filename=OutPut_Filename)
else:
    workbook = Workbook()

sheet = workbook.active
sheet.cell(1, 1).value = 'question'
# sheet.cell(1, 2).value = 'answer'
sheet.column_dimensions['A'].width = 25
sheet.column_dimensions['B'].width = 35

def setup_api(browser_driver,api_key, app):
    api_button = browser_driver.find_element(By.CSS_SELECTOR, 'nav a:nth-child(4)')
    api_button.click()
    time.sleep(sleep_time)

    app.top_window().send_keystrokes("{TAB 3}")
    time.sleep(sleep_time)
    app.top_window().send_keystrokes(api_key)
    time.sleep(sleep_time)
    app.top_window().send_keystrokes("{TAB 3}")
    time.sleep(sleep_time)
    app.top_window().send_keystrokes("{ENTER 1}")
    time.sleep(sleep_time)
    app.top_window().send_keystrokes("{TAB 10}")
    time.sleep(sleep_time)

    # # api_form = WebDriverWait(browser_driver, request_time_out_seconds).until(EC.visibility_of_element_located(
    # #     (By.CSS_SELECTOR, '.prose')))
    # api_form = browser_driver.find_element(By.CSS_SELECTOR, '.prose')
    # print('---- api_form.text: ', api_form.text)
    # api_input = WebDriverWait(browser_driver, request_time_out_seconds).until(EC.visibility_of_element_located(
    #     (By.TAG_NAME, 'input')))
    # api_input.send_keystrokes(api_key)
    # time.sleep(sleep_time)
    # time.sleep(20)
    # api_save_button = browser_driver.find_elements(By.CSS_SELECTOR, '.btn-primary div')
    # api_save_button.click()
    # time.sleep(sleep_time)
    

def open_chatgpt():
    global URL, sheet
    sheet.cell(1, 2).value = 'answer'
    my_chrome_options = webdriver.ChromeOptions()
    my_chrome_options.add_argument('--ignore-certificate-errors')
    my_chrome_options.add_argument('--ignore-ssl-errors')
    my_chrome_options.add_argument('--start-maximized')
    # my_chrome_options.add_argument('--ignore-certificate-errors')
    my_chrome_options.add_argument('--ignore-certificate-errors-spki-list')
    if BrowserHide:
        my_chrome_options.add_argument('--headless')
    URL = 'https://fastgpt.app/'
    try:
        if selenium_connect_method == 1:
            if chromedriver_path != None:
                browser_driver = webdriver.Chrome(
                    chromedriver_path, options=my_chrome_options)
            else:
                browser_driver = webdriver.Chrome(options=my_chrome_options)
        elif selenium_connect_method == 2:
            from webdriver_manager.chrome import ChromeDriverManager
            from selenium.webdriver.chrome.service import Service
            browser_driver = webdriver.Chrome(service=Service(
                ChromeDriverManager().install()), options=my_chrome_options)
        elif selenium_connect_method == 3:
            browser_driver = webdriver.Remote(
                command_executor=liara_chrome_app_url, options=my_chrome_options)
        browser_driver.get(URL)
        time.sleep(sleep_time)
    except Exception as error:
        print('---------- Exception html_format : --------\n', error)
    else:  # if try part was successful
        browser_driver.execute_script(
            'document.title = "%s"' % BrowserTitleToFind)
        time.sleep(sleep_time)

        app = Application().connect(title_re=BrowserTitleToFind)
        time.sleep(sleep_time)

        if api_key != '':
            print('-- setting up api_key --')
            setup_api(browser_driver,api_key, app)

        your_question = ''
        while your_question != 'stop':
            numrows = sheet.max_row
            your_question = input('ask some question from chat-GPT: ')
            if your_question == 'stop':
                break

            # method 1
            app.top_window().send_keystrokes(your_question)
            time.sleep(sleep_time)
            app.top_window().send_keystrokes("{ENTER 1}")
            time.sleep(sleep_time)

            # method 2
            # text_area = browser_driver.find_element(By.CSS_SELECTOR, 'textarea')
            # time.sleep(sleep_time)
            # your_question = input('ask some question from chat-GPT: ')
            # text_area.__setattr__('value',your_question)
            # time.sleep(sleep_time)
            # send_button = browser_driver.find_element(By.CSS_SELECTOR, 'button.absolute')
            # send_button.click()
            # time.sleep(sleep_time)

            WebDriverWait(browser_driver, request_time_out_seconds).until(EC.visibility_of_all_elements_located(
                (By.CSS_SELECTOR, '.items-start.gap-4')))
            elements = browser_driver.find_elements(By.CSS_SELECTOR, '.items-start.gap-4')
            print(elements[-1].text)

        elements = browser_driver.find_elements(By.CSS_SELECTOR, '.items-start.gap-4')
        time.sleep(sleep_time)
        numrows = 0
        for element_index in range(len(elements)):
            if element_index % 2 == 0:
                numrows = sheet.max_row
                # print('-numrows:', numrows)
                sheet.cell(numrows+1, 1).value = elements[element_index].text
            else:
                sheet.cell(numrows+1, 2).value = elements[element_index].text
            print('=== element number %i, numrow is: %i, text is: %s' %
                  (element_index+1, numrows, elements[element_index].text))
        try:
            workbook.save(OutPut_Filename)
        except Exception as error:
            if debug:
                print('-- error on save to excel. error message is:\n',error)
        # browser_driver.quit()


open_chatgpt()
