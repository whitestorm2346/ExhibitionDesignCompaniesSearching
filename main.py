from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from colorama import init, Fore, Back, Style
import chromedriver_autoinstaller
import pandas as pd
import re

email_regex = r"""
    (?:[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+)*
    |"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]
    |\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")
    @
    (?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+
    [a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?
"""

def scroll_and_load(driver, scroll_pause_time=1) -> int:
    scroll_element = driver.find_elements(By.CLASS_NAME, 'ecceSd')[1]

    last_height = driver.execute_script("return arguments[0].scrollHeight", scroll_element)
    
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scroll_element)
    sleep(scroll_pause_time)

    new_height = driver.execute_script("return arguments[0].scrollHeight", scroll_element)
        
    if new_height == last_height: # 代表不能再往下滾動了
        return 1
    
    return 0

def extract_email(email):
    
    extracted_email = re.findall(email_regex, email, re.VERBOSE)

    if extracted_email == []:
        return ""
    else:
        return extracted_email[0]
    
def get_pseudo_content(driver, element, pseudo):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            return driver.execute_script("""
                var element = arguments[0];
                var pseudoElement = arguments[1];
                var content = window.getComputedStyle(element, pseudoElement).getPropertyValue('content');
                return content;
                """, element, pseudo)
        except StaleElementReferenceException:
            sleep(1)
    return None

def scan_page_for_email(driver, url):
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    driver.implicitly_wait(3)

    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')
    page_text = soup.get_text(separator="\n")

    emails = []
    emails.extend(re.findall(email_regex, page_text, re.VERBOSE))

    if emails == []:
        elements = driver.find_elements(By.CSS_SELECTOR, '*')

        for element in elements:
            for pseudo in ['::before', '::after']:
                for i in range(3):
                    try:
                        pseudo_content = driver.execute_script("""
                            var element = arguments[0];
                            var pseudoElement = arguments[1];
                            var content = window.getComputedStyle(element, pseudoElement).getPropertyValue('content');
                            return content;
                            """, element, pseudo)
                        break
                    except StaleElementReferenceException as e:
                        pass
                    
                if pseudo_content and pseudo_content != 'none':
                    emails.extend(re.findall(email_regex, pseudo_content, re.VERBOSE))
                

    if emails == []:
        emails = driver.find_elements(By.XPATH, '//a[contains(@href, "mailto:")]')
        emails = [email.get_attribute('href') for email in emails]

    return emails

def get_emails(driver, url):
    emails = []

    if 'www.facebook.com' in url:
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        driver.implicitly_wait(3)

        try:
            close_btn = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[aria-label="關閉"]')))
            close_btn.click()
        except TimeoutException as e:
            pass

        try:
            infos = driver.find_elements(By.CLASS_NAME, 'xieb3on')[0]
            emails = re.findall(email_regex, infos.text, re.VERBOSE)
        except Exception as e:
            pass
    else:
        # find email in main page

        emails = scan_page_for_email(driver, url)

        # find email in contact page

        if emails == []:
            contact_links = driver.find_elements(By.XPATH, '//a[contains(@href, "contact_us") or contains(@href, "info") or contains(@href, "contact") or contains(@href, "contacts")]')

            for contact_link in contact_links:
                try:
                    link = contact_link.get_attribute('href')
                except StaleElementReferenceException as e:
                    continue

                emails.extend(scan_page_for_email(driver, link))

    emails = list(set([extract_email(email) for email in emails]))

    return emails

def print_company_info(company):
    emails = '[' + ", ".join(company['E-mail']) + ']'

    print(Style.BRIGHT + '<company> ' + Fore.LIGHTWHITE_EX + company['公司名稱'])
    print(Style.BRIGHT + '<email>   ' + Fore.LIGHTYELLOW_EX + emails)
    print(Style.BRIGHT + '<phone>   ' + Fore.LIGHTGREEN_EX + company['電話'])
    print(Style.BRIGHT + '<website> ' + Fore.LIGHTCYAN_EX + company['網站'])
    print(Style.BRIGHT + '<score>   ' + Fore.YELLOW + company['Google評分'])
    print(Style.BRIGHT + '<address> ' + Fore.MAGENTA + company['地址'])


init(autoreset=True)

keyword = "展覽設計公司"
map_search_url = f"https://www.google.com.tw/maps/search/{keyword}"

chrome_option = chromeOptions()
chrome_option.add_argument('--log-level=3')
chrome_option.add_argument('--headless') 
# chrome_option.add_argument('--start-maximized') 
chrome_option.add_argument("--retry-on-network-failure")
chromedriver_autoinstaller.install()
driver = webdriver.Chrome(options=chrome_option)

driver.get(map_search_url)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

companies = []

while len(companies) <= 200:
    companies_info = driver.find_elements(By.CLASS_NAME, 'Nv2PK')

    for company_info in companies_info:
        # Advertise Message
        try: 
            sponser = company_info.find_element(By.CLASS_NAME, 'kpih0e')
            continue
        except NoSuchElementException as e:
            pass

        company_info = company_info.find_element(By.CLASS_NAME, 'lI9IFe')

        try:
            url = company_info.find_element(By.TAG_NAME, 'a').get_attribute('href')
        except NoSuchElementException as e:
            continue # if the company does not have the website

        name = company_info.find_element(By.CLASS_NAME, 'qBF1Pd').text

        try:
            score = company_info.find_element(By.CLASS_NAME, 'MW4etd').text
        except NoSuchElementException as e:
            score = 'no comments'

        try:
            phone = company_info.find_element(By.CLASS_NAME, 'UsdlK').text
        except NoSuchElementException as e:
            phone = 'null'

        address_info = company_info.find_elements(By.CLASS_NAME, 'W4Efsd')[2].text.split('·')

        # company's type
        classification = address_info[0]
        address = "".join(address_info[-1].split())

        company = {
            "公司名稱": name,
            "E-mail": "",
            "網站": url,
            "電話": phone,
            "Google評分": score,
            "地址": address
        }

        if not company in companies:
            companies.append(company)

    if scroll_and_load(driver):
        break

results = []

for company in companies:
    company['E-mail'] = get_emails(driver, company['網站'])

    print_company_info(company)
    print('\n')

    if company['E-mail'] != []:
        results.append(company)


driver.close()

df = pd.DataFrame(results)
df.to_excel('展覽設計公司.xlsx', index=False)