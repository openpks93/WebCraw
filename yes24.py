from selenium import webdriver

print("set url : " + 'https://scm.yes24.com/Login/LogOn?ReturnUrl=/Home/Index')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(20)
chrome_driver.get('https://scm.yes24.com/Login/LogOn?ReturnUrl=/Home/Index')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('UserName').send_keys('tg7616')
chrome_driver.find_element_by_id('Password').send_keys('p@ssw0rd6806!!')
chrome_driver.find_element_by_xpath('//*[@id="LogOnForm"]/fieldset/div[1]/div[2]/div').click()
#조회 클릭

# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

# 크롤링 element

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
with open('yes24.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('yes24.json',  encoding='UTF8', orient='index').to_excel('yes24.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

