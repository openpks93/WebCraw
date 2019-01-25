from selenium import webdriver

print("set url : " + 'http://soffice.11st.co.kr/Index.tmall')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(5)
chrome_driver.get('http://soffice.11st.co.kr/Index.tmall')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('loginName').send_keys('tg7616')
chrome_driver.find_element_by_id('passWord').send_keys('aseman6806!')
chrome_driver.find_element_by_class_name('btn_login').click()

#판매정산형황
chrome_driver.find_element_by_xpath('//*[@id="35936"]/a').click()
#월별 클릭이 안됨
chrome_driver.find_elements_by_xpath('//*[@id="searchPrdDateCondition"]').click()
#검색
chrome_driver.find_element_by_xpath('//*[@id="btnSearch"]/span/span').click()
print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

Order_Number = chrome_driver.find_elements_by_css_selector("#row0dataGrid > div:nth-child(2) > div")
shipping_number = chrome_driver.find_elements_by_css_selector("#row0dataGrid > div:nth-child(4) > div")
Order_Status = chrome_driver.find_elements_by_css_selector("#row0dataGrid > div:nth-child(5) > div")
Buyer_name = chrome_driver.find_elements_by_css_selector("#row0dataGrid > div:nth-child(6) > div")

# Json
import json
from collections import OrderedDict
LENGTH = min(len(Order_Number), len(shipping_number), len(Order_Status))
products = OrderedDict()
print("processing data to json")

for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'Item_Order_Number': Order_Number[idx].text,
        'shipping_number': shipping_number[idx].text,
        'Order_Status': Order_Status[idx].text,
        'Buyer_name': Buyer_name[idx].text
    }

# json 저장
with open('eleventh.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('eleventh.json',  encoding='UTF8').to_excel('eleventh.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

