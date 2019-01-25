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

chrome_driver.implicitly_wait(20)
chrome_driver.get('https://sell.smartstore.naver.com/#/naverpay/manage/order')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('loginId').send_keys('tg7616')
chrome_driver.find_element_by_id('loginPassword').send_keys('p@ssw0rd6806!')
chrome_driver.find_element_by_id('loginButton').click()
print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")
Item_Order_Number = chrome_driver.find_elements_by_css_selector("body > div.npay_content._root > div.napy_sub_content > div:nth-child(2) > div.npay_grid_area > div.grid._detailGrid._grid_container.uio_grid > div._inflexible_area.inflexible_area > div._body.body > div._table_container > table > tbody > tr > td:nth-child(2)")
Order_Numver = chrome_driver.find_elements_by_css_selector("body > div.npay_content._root > div.napy_sub_content > div:nth-child(2) > div.npay_grid_area > div.grid._detailGrid._grid_container.uio_grid > div._inflexible_area.inflexible_area > div._body.body > div._table_container > table > tbody > tr > td:nth-child(3)")
Order_Date = chrome_driver.find_elements_by_css_selector("body > div.npay_content._root > div.napy_sub_content > div:nth-child(2) > div.npay_grid_area > div.grid._detailGrid._grid_container.uio_grid > div._flexible_area.flexible_area > div._body.body > div._table_container > table > tbody > tr > td:nth-child(1)")

print(Item_Order_Number)
print(Order_Numver)
print(Order_Date)
# Json
import json
from collections import OrderedDict
LENGTH = min(len(Item_Order_Number), len(Order_Numver), len(Order_Date))
products = OrderedDict()
print("processing data to json")

for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'Item_Order_Number': Item_Order_Number[idx].text,
        'Order_Number': Order_Numver[idx].text,
        'Order_Date': Order_Date[idx].text
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

