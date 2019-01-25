from selenium import webdriver
#url = r"https://sell.smartstore.naver.com/#/products/origin-list"
print("set url : " + 'http://www.dekorea.co.kr/shop/member.html?type=login')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(30)
chrome_driver.get('http://www.dekorea.co.kr/shop/member.html?type=login')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_elements_by_name('id').send_keys('dekorea78')
chrome_driver.find_element_by_id('passwd').send_keys('aseman6806!')
chrome_driver.find_element_by_xpath('//*[@id="mk_center"]/table/tbody/tr[4]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[5]/a/img').click()

print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

# Json
import json
from collections import OrderedDict
LENGTH = min()
products = OrderedDict()
print("processing data to json")
for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
    }

# json 저장
with open('dekorea.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('dekorea.json',  encoding='UTF8', orient='index').to_excel('dekorea.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

