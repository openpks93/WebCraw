from selenium import webdriver

print("set url : " + 'http://seller.nhmarket.kr/mallinmall/order/order_list.asp?menu=delivery')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(3)
chrome_driver.get('http://seller.nhmarket.kr/mallinmall/order/order_list.asp?menu=delivery')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('id').send_keys('tg7616')
chrome_driver.find_element_by_id('pass').send_keys('p@ssw0rd6806!')
chrome_driver.find_element_by_xpath('//*[@id="login_wrap"]/div[4]/img').click()
#배달 조회를 위한 것 페이지로 안넘어 가져서 패스를 주고 넘어가게 한다 하하하하
chrome_driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[1]/tbody/tr/td[5]/a/img').click()
# 발송요청 버튼
chrome_driver.find_element_by_xpath('/html/body/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/div/input[2]').click()
#발송 완료 버튼
chrome_driver.find_element_by_xpath('/html/body/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/div/input[3]').click()
#조회
chrome_driver.find_element_by_xpath('/html/body/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form[1]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[3]/input[2]').click()

print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

Item_number = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form:nth-child(6) > table > tbody > tr > td:nth-child(2) > div > u")
product_pr = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form:nth-child(6) > table > tbody > tr > td > table > tbody > tr:nth-child(3) > td:nth-child(4)")
shipping_fee = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form > table > tbody > tr > td:nth-child(4)")
Shipping_Method = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form > table > tbody > tr > td:nth-child(5)")
buyer = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form > table > tbody > tr > td:nth-child(6)")
recipient = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form > table > tbody > tr > td:nth-child(7)")
order_date = chrome_driver.find_elements_by_css_selector("body > div > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > table > tbody > tr > td > form > table > tbody > tr > td:nth-child(8)")
# Json
import json
from collections import OrderedDict
LENGTH = min(len(Item_number), len(product_pr), len(shipping_fee), len(Shipping_Method), len(buyer), len(recipient), len(order_date))
products = OrderedDict()
print("processing data to json")
for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'Item_number': Item_number[idx].text,
        'product_pr': product_pr[idx].text,
        'shipping_fee': shipping_fee[idx].text,
        'Shipping_Method': Shipping_Method[idx].text,
        'buyer': buyer[idx].text,
        'recipient': recipient[idx].text,
        'order_date': order_date[idx].text
    }
# json 저장
with open('nh.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('nh.json',  encoding='UTF8', orient='index').to_excel('nh.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

