from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import chromedriver_autoinstaller
import pandas as pd
import time
import re

taiwan_address_pattern_advanced = re.compile(
    r"""
    (?P<county>[^市縣]{1,3}(?:市|縣))\s*           # 縣市，如「台北市」
    (?P<district>[^鄉鎮市區]{1,3}[鄉鎮市區])?\s*    # 鄉鎮市區，如「中正區」
    (?P<street>[^路街段巷弄號樓F\d]{1,}(?:路|街|段))?\s* # 街道和段，如「中山北路二段」
    (?P<alley>[^弄號樓F\d]{1,}(?:巷))?\s*         # 巷，如「34巷」
    (?P<lane>[^號樓F\d]{1,}(?:弄))?\s*            # 弄，如「8弄」
    (?P<number>\d+(?:-\d+)?號(?:[A-Za-z])?)?\s*    # 號碼，支持19-6號、12號A
    (?P<floor>(?:[之]*\d*[樓Ff](?:之\d+)?))?\s*   # 樓層，支持「3樓」、「10F」、「3樓之1」
    (?:\((?P<additional_info>[^)]+)\))?           # 附加信息，如「南港軟體工業園區」
    """, re.VERBOSE
)
phone_fax_pattern = re.compile(
    r"""
    (?:(?:\b(?:電話|TEL|Tel|tel)\b[:\s]*)|    # 電話的上下文
    (?:\b(?:傳真|FAX|Fax|fax)\b[:\s]*))?      # 傳真的上下文
    (
        (?:\+886\s?)?                           # 國際區碼
        (?:\(?(0\d{1,4})\)?[\s\-]?)?            # 區碼
        \d{1,4}[\s\-]?\d{1,4}(?:[\s\-]?\d{1,4})? # 號碼段
        |                                       # 或者
        09\d{2}[\s\-]?\d{3}[\s\-]?\d{3}         # 手機號碼
    )
    """, re.VERBOSE
)
simple_email_pattern = re.compile(
    r"""
    # 本地部分
    [a-zA-Z0-9._%+-]+
    @
    # 域名部分
    [a-zA-Z0-9.-]+\.[a-zA-Z]{2,}
    """, re.VERBOSE
)

keyword = "展覽設計公司"
search_url = f"https://www.google.com/search?q={keyword}"

chrome_option = chromeOptions()
chrome_option.add_argument('--log-level=3')
chrome_option.add_argument('--headless') 
chromedriver_autoinstaller.install()
driver = webdriver.Chrome(options=chrome_option)

driver.get(search_url)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

page_content = driver.page_source
soup = BeautifulSoup(page_content, 'html.parser')

search_results = soup.find_all('div', class_='g') 
data = []

for result in search_results:
    company_name = result.find('span', class_='VuuXrf')
    link = result.find('a')['href']

    print(f'<company> {company_name}')
    print(f'<link> {link}')

    driver.get(link)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

    company_page = driver.page_source
    company_soup = BeautifulSoup(company_page, 'html.parser')
    company_text_content = company_soup.get_text(separator='\n')

    addresses = taiwan_address_pattern_advanced.findall(company_text_content)
    phones = phone_fax_pattern.findall(company_text_content)
    emails = simple_email_pattern.findall(company_text_content)

    addresses = list(set(addresses))
    phones = list(set(phones))
    emails = list(set(emails))

    company_info = {
        "公司名稱": company_name.text if company_name else "",
        "網址": link,
        "地址": addresses, 
        "電話": phones,  
        "Email": emails,  
    }

    data.append(company_info)

driver.quit()

df = pd.DataFrame(data)
df.to_excel('展覽設計公司.xlsx', index=False)



# num_pages = 5  # 需要抓取的页面数量
# for page in range(num_pages):
#     print(f"正在处理第 {page + 1} 页...")
#     # links = extract_links()
#     # all_links.update(links)
#     try:
#         # 点击下一页
#         next_button = driver.find_element(By.ID, 'pnnext')
#         next_button.click()
#         time.sleep(2)  # 等待页面加载
#     except Exception as e:
#         print("没有更多页面了。")
#         break