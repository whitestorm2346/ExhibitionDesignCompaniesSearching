from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from colorama import init, Fore, Back, Style
import chromedriver_autoinstaller
import pandas as pd


def scroll_and_load(driver, scroll_pause_time=2) -> int:
    scroll_element = driver.find_elements(By.CLASS_NAME, 'ecceSd')[1]

    last_height = driver.execute_script("return arguments[0].scrollHeight", scroll_element)
    
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scroll_element)
    sleep(scroll_pause_time)

    new_height = driver.execute_script("return arguments[0].scrollHeight", scroll_element)
        
    if new_height == last_height: # 代表不能再往下滾動了
        return 1
    
    return 0


init(autoreset=True)

keyword = "展覽設計公司"
map_search_url = f"https://www.google.com.tw/maps/search/{keyword}"

chrome_option = chromeOptions()
chrome_option.add_argument('--log-level=3')
chrome_option.add_argument('--headless') 
# chrome_option.add_argument('--start-maximized') 
chromedriver_autoinstaller.install()
driver = webdriver.Chrome(options=chrome_option)

driver.get(map_search_url)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

results = []

while len(results) <= 20:
    companies_info = driver.find_elements(By.CLASS_NAME, 'Nv2PK')

    for company_info in companies_info:
        try: 
            sponser = company_info.find_element(By.CLASS_NAME, 'kpih0e')
            continue
        except NoSuchElementException as e:
            pass

        company_info = company_info.find_element(By.CLASS_NAME, 'lI9IFe')

        try:
            url = company_info.find_element(By.TAG_NAME, 'a').get_attribute('href')
        except NoSuchElementException as e:
            url = 'null'

        name = company_info.find_element(By.CLASS_NAME, 'qBF1Pd').text

        try:
            score = company_info.find_element(By.CLASS_NAME, 'MW4etd').text
        except NoSuchElementException as e:
            score = 'no comments'

        address_info = company_info.find_elements(By.CLASS_NAME, 'W4Efsd')[2].text.split('·')

        classification = address_info[0]
        address = "".join(address_info[-1].split())

        try:
            phone = company_info.find_element(By.CLASS_NAME, 'UsdlK').text
        except NoSuchElementException as e:
            phone = 'null'

        company = {
            "公司名稱": name,
            "網站": url,
            "Google評分": score,
            "地址": address,
            "電話": phone
        }

        print(Style.BRIGHT + '<company> ' + Fore.LIGHTWHITE_EX + name)
        print(Style.BRIGHT + '<website> ' + Fore.LIGHTCYAN_EX + url)
        print(Style.BRIGHT + '<score>   ' + Fore.YELLOW + score)
        print(Style.BRIGHT + '<address> ' + Fore.MAGENTA + address)
        print(Style.BRIGHT + '<phone>   ' + Fore.LIGHTGREEN_EX + phone)
        print('\n')

        if not company in results:
            results.append(company)

    if scroll_and_load(driver, 2):
        break


driver.close()

df = pd.DataFrame(results)
df.to_excel('展覽設計公司.xlsx', index=False)