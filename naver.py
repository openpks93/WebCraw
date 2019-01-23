from selenium import webdriver
#url = r"https://sell.smartstore.naver.com/#/products/origin-list"
print("set url : " + 'https://sell.smartstore.naver.com/#/products/origin-list')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")


# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(3)
chrome_driver.get('https://sell.smartstore.naver.com/#/products/origin-list')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('loginId').send_keys('tg7616')
chrome_driver.find_element_by_id('loginPassword').send_keys('p@ssw0rd6806!')
chrome_driver.find_element_by_id('loginButton').click()
print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")
Item_number = chrome_driver.find_elements_by_css_selector("div.ag-fresh > div.ag-bl-normal > div.ag-bl-center-row > div.ag-bl-normal-center > div.ag-bl > div.ag-bl-center > div.ag-root > div.ag-body > div.ag-pinned-left-cols-viewport > div.ag-pinned-left-cols-container > div.ag-row > div.ag-cell-no-focus > a.text-info")
product_name = chrome_driver.find_elements_by_css_selector("div.ag-body-viewport-wrapper > div.ag-body-viewport > div.ag-body-container > div.ag-row > div.ag-cell-ncp-editable")

# Item_Order_Number = chrome_driver.find_elements_by_css_selector('div._table_container > td.ellipsis > a._click')
# Order_Number = chrome_driver.find_elements_by_css_selector('div._body > div._table_container > table > tbody > tr > td.ellipsis')


print(Item_number)
print(product_name)

# Json
import json
from collections import OrderedDict
LENGTH = min(len(Item_number), len(product_name))
products = OrderedDict()
print("processing data to json")

for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'Item_Order_Number': Item_number[idx].text,
        'Order_Number': product_name[idx].text
    }

# json 저장
with open('naver.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('naver.json',  encoding='UTF8').to_excel('naver.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

