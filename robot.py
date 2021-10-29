from pywinauto import Application
from pywinauto import clipboard
import time

class CrmAppAgent:
    def __init__(self):
        window_title_regular_expression = ".*data.*"
        excel_app = Application(backend="uia").connect(title_re=window_title_regular_expression)
        excel_window = excel_app.window(title_re=window_title_regular_expression)
        self.excel_app = excel_app
        self.excel_window = excel_window

    def deselect(self):
        excel_window = self.excel_window
        excel_window.type_keys('{ESC}')

    def move_cursor_to_top_left_corner(self):
        excel_window = self.excel_window
        excel_window.type_keys('^{HOME}')

    def move_cursor_down(self):
        excel_window = self.excel_window
        excel_window.type_keys('{DOWN}')

    def move_cursor_to_first_person_cell(self):
        excel_window = self.excel_window
        self.move_cursor_to_top_left_corner()
        self.move_cursor_down()

    def read_cell_contents(self):
        excel_window = self.excel_window
        excel_window.type_keys('^c')
        time.sleep(0.1)
        cell_data = clipboard.GetData()
        result = cell_data.rstrip()
        return result

    def move_cursor_right(self):
        excel_window = self.excel_window
        excel_window.type_keys('{RIGHT}')

    def move_cursor_to_first_left_cell(self):
        excel_window = self.excel_window
        excel_window.type_keys('{HOME}')

    def read_person(self):
        excel_window = self.excel_window
        last_name = self.read_cell_contents()

        if not last_name:
            return None

        self.move_cursor_right()
        first_name = self.read_cell_contents()

        self.move_cursor_right()
        middle_name = self.read_cell_contents()

        self.move_cursor_right()
        birthday = self.read_cell_contents()

        self.move_cursor_right()
        passport = self.read_cell_contents()

        self.move_cursor_down()
        self.move_cursor_to_first_left_cell()

        return {
            "last_name": last_name,
            "first_name": first_name,
            "middle_name": middle_name,
            "birthday": birthday,
            "passport": passport
        }

    def read_persons(self):
        excel_window = self.excel_window
        result = []
        self.deselect()
        self.move_cursor_to_first_person_cell()
        while True:
            person = self.read_person()
            if person:
                result.append(person)
            else:
                break

        return result
    
    def move_cursor_to_first_inn(self):
        excel_window = self.excel_window
        self.move_cursor_to_top_left_corner()
        self.move_cursor_down()

        for i in range(5):
            self.move_cursor_right()
        
    def fill_inns(self, inns):
        excel_window = self.excel_window
        self.move_cursor_to_first_inn()
        for inn in inns:
            excel_window.type_keys(inn)
            self.move_cursor_down()

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os

class InnAppAgent:
    def __init__(self):        
        browser = self.get_browser()
        self.browser = browser
        
        browser.implicitly_wait(10)

        SERVICE_URL = 'https://service.nalog.ru/inn.do'
        browser.get(SERVICE_URL)

        accept_terms_and_conditions_page_shown = browser.current_url != SERVICE_URL

        if accept_terms_and_conditions_page_shown:
            self.accept_terms_and_conditions()
            
    def get_browser(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

        current_folder = os.getcwd()
        web_driver_executable_name = "chromedriver.exe"
        web_driver_executable_path = "{}\\{}".format(current_folder, web_driver_executable_name)
        result = webdriver.Chrome(executable_path=web_driver_executable_path, chrome_options=options)
        return result

    def accept_terms_and_conditions(self):
        browser = self.browser
        browser.find_element(By.XPATH, '//a[@class="checkbox checkbox-off"]').click()
        browser.find_element(By.XPATH, '//button[@id="btnContinue"]').click()

    def fill_person_data(self, person):
        browser = self.browser
        input_data = {
            "fam": person["last_name"],
            "nam": person["first_name"],
            "otch": person["middle_name"],
            "bdate": person["birthday"],
            "docno": person["passport"]
        }

        for element_id, input_value in input_data.items():
            element = browser.find_element(By.ID, element_id)
            element.clear()
            for symbol in input_value:
                element.send_keys(symbol)
                time.sleep(0.1)

    def submit_data(self):
        browser = self.browser
        browser.find_element(By.ID, 'btn_send').click()

    def read_inn(self):
        browser = self.browser
        previous_inn_element = browser.find_element(By.ID, "resultInn")
        previous_inn = previous_inn_element.text
        WebDriverWait(driver=browser, timeout=10, poll_frequency=1).until(lambda drv: drv.find_element(By.ID, "resultInn").text != previous_inn)

        inn_element = browser.find_element(By.ID, "resultInn")
        result = inn_element.text

        return result

    def submit_data_and_read_inn(self):
        browser = self.browser
        self.submit_data()

        return self.read_inn()

    def find_inn(self, person):
        browser = self.browser
        self.fill_person_data(person=person)

        result = self.submit_data_and_read_inn()

        return result

    def find_inns(self, persons):    
        browser = self.browser

        return [self.find_inn(person=person) for person in persons]

class EnrichPersonsWithInnsScenario:
    def __init__(self, crm_app_agent, inn_app_agent):
        self.crm_app_agent = crm_app_agent
        self.inn_app_agent = inn_app_agent
        
    def run(self):
        crm_app_agent = self.crm_app_agent

        persons = crm_app_agent.read_persons()

        inn_app_agent = self.inn_app_agent
        inns = inn_app_agent.find_inns(persons=persons)

        crm_app_agent.fill_inns(inns=inns)

crm_app_agent = CrmAppAgent()
inn_app_agent = InnAppAgent()
enrich_persons_with_inns_scenario = EnrichPersonsWithInnsScenario(crm_app_agent=crm_app_agent, inn_app_agent=inn_app_agent)
enrich_persons_with_inns_scenario.run()

