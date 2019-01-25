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
chrome_driver.implicitly_wait(30)
chrome_driver.get('https://spc.ticketmonster.co.kr/member/login?return_url=%2F')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('form_id').send_keys('태광인테리어필름')
chrome_driver.find_element_by_id('form_password').send_keys('p@ssw0rd6806!')
chrome_driver.find_element_by_xpath('//*[@id="content"]/div[1]/form/ul/fieldset/li[5]/button').click()
# 비번 다음에 바꾸기 버튼 클릭
chrome_driver.implicitly_wait(5)
chrome_driver.find_element_by_class_name('btn_delay_modify').click()
# 공지창 제거 버튼
chrome_driver.find_element_by_xpath('//*[@id="notice5236"]/div[3]/button').click()
chrome_driver.find_element_by_xpath('//*[@id="notice5234"]/div[3]/button').click()
# 배송 관리
chrome_driver.find_element_by_xpath('//*[@id="gnb"]/ul/li[4]/a').click()
chrome_driver.find_element_by_xpath('//*[@id="btnSearch"]').click()

print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")
shipping_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(3)")
deal_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(4)")
order_Number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(5)")
courier_company = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(6)")
waybill_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(7) > a")
shipping_date = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(8)")
Delayed_due_date = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(9)")
reason_for_delaying = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(10)")
Details_of_delay_reason = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(11)")
Delayed_reporting_date = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(12)")
When_to_register = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(13)")
order_name = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(14)")
user_id = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(15)")
order_phone_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(16)")
item_name = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(17)")
option_name = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(18)")
unit_price = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(19)")
purchase_quantity = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(20)")
price = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(21)")
payment_completion_date = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(22)")
Additional_phrases = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(23)")
payee_name = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(24)")
Resident_registration_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(25)")
Personal_clearance_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(26)")
payee_phone_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(27)")
address = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(28)")
post_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(29)")
shipping_notes = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(30)")
Safe_number_Effective_period = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(31)")
shipping_status = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(32)")
final_processing_date = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(33)")
partner_code1 = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(34)")
partner_code2 = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(35)")
partner_code3 = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(36)")
partner_code4 = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(37)")
partner_code5 = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(38)")
Additional_Invoice = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(39)")
Option_number = chrome_driver.find_elements_by_css_selector("#__grid_DeliveryGrid > div.objbox > table > tbody > tr > td:nth-child(40)")

# Json
import json
from collections import OrderedDict
LENGTH = min(len(shipping_number), len(deal_number), len(order_Number), len(courier_company), len(waybill_number), len(shipping_date), len(Delayed_due_date), len(reason_for_delaying), len(Details_of_delay_reason), len(Delayed_reporting_date), len(When_to_register), len(order_name), len(user_id), len(order_phone_number), len(item_name), len(option_name), len(unit_price), len(purchase_quantity), len(price), len(payment_completion_date), len(Additional_phrases), len(payee_name), len(Resident_registration_number), len(Personal_clearance_number), len(payee_phone_number), len(address), len(post_number), len(shipping_notes), len(Safe_number_Effective_period), len(shipping_status), len(final_processing_date), len(partner_code1), len(partner_code2), len(partner_code3), len(partner_code4), len(partner_code5), len(Additional_Invoice), len(Option_number))
products = OrderedDict()
print("processing data to json")
for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'shipping_number': shipping_number[idx].text,
        'deal_number': deal_number[idx].text,
        'order_Number': order_Number[idx].text,
        'courier_company': courier_company[idx].text,
        'waybill_number': waybill_number[idx].text,
        'shipping_date' : shipping_date[idx].text,
        'Delayed_due_date': Delayed_due_date[idx].text,
        'reason_for_delaying': reason_for_delaying[idx].text,
        'Details_of_delay_reason': Details_of_delay_reason[idx].text,
        'Delayed_reporting_date': Delayed_reporting_date[idx].text,
        'When_to_register': When_to_register[idx].text,
        'order_name': order_name[idx].text,
        'user_id': user_id[idx].text,
        'order_phone_number': order_phone_number[idx].text,
        'item_name': item_name[idx].text,
        'option_name': option_name[idx].text,
        'unit_price': unit_price[idx].text,
        'purchase_quantity': purchase_quantity[idx].text,
        'price': price[idx].text,
        'payment_completion_date': payment_completion_date[idx].text,
        'Additional_phrases' : Additional_phrases[idx].text,
        'payee_name' : payee_name[idx].text,
        'Resident_registration_number': Resident_registration_number[idx].text,
        'Personal_clearance_number': Personal_clearance_number[idx].text,
        'payee_phone_number': payee_phone_number[idx].text,
        'address': address[idx].text,
        'post_number': post_number[idx].text,
        'shipping_notes': shipping_notes[idx].text,
        'Safe_number_Effective_period': Safe_number_Effective_period[idx].text,
        'shipping_status': shipping_status[idx].text,
        'final_processing_date': final_processing_date[idx].text,
        'partner_code1': partner_code1[idx].text,
        'partner_code2': partner_code2[idx].text,
        'partner_code3': partner_code3[idx].text,
        'partner_code4': partner_code4[idx].text,
        'partner_code5': partner_code5[idx].text,
        'Additional_Invoice': Additional_Invoice[idx].text,
        'Option_number': Option_number[idx].text
    }

# json 저장
with open('tmon.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('tmon.json',  encoding='UTF8', orient='index').to_excel('tmon.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

