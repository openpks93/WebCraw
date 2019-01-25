from selenium import webdriver

print("set url : " + 'https://www.esmplus.com/Member/SignIn/LogOn')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(20)
chrome_driver.get('https://www.esmplus.com/Member/SignIn/LogOn')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_xpath('//*[@id="captchaLogin"]/div[1]/div/div/label[1]').click()
chrome_driver.find_element_by_id('SiteId').send_keys('tg7616')
chrome_driver.find_element_by_id('SitePassword').send_keys('aseman6806!')
chrome_driver.find_element_by_xpath('//*[@id="btnSiteLogOn"]/img').click()

#
chrome_driver.implicitly_wait(40)
chrome_driver.find_element_by_xpath('//*[@id="TransDueDateToday"]').click()

# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

# 크롤링 element
user_id = chrome_driver.find_elements_by_css_selector('#dataGrid > table > tbody > tr > td.span_left')
Date_decision = chrome_driver.find_elements_by_css_selector('#dataGrid > table > tbody > tr > td:nth-child(3)')
Settlement_status = chrome_driver.find_elements_by_css_selector('#dataGrid > table > tbody > tr > td:nth-child(4)')


print(user_id)
print(Date_decision)
# Json
import json
from collections import OrderedDict
LENGTH = min(len(user_id), len(Date_decision), len(Settlement_status))
products = OrderedDict()
print("processing data to json")

for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'user_id': user_id[idx].text,
        'Date_decision': Date_decision[idx].text,
        'Settlement_status': Settlement_status[idx]
    }

# json 저장
with open('auction.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('auction.json',  encoding='UTF8', orient='index').to_excel('auction.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

