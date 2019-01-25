from selenium import webdriver
driver = webdriver.Chrome(r'C:\Program Files (x86)\Google\Chrome\chromedriver.exe')

driver.implicitly_wait(10)

driver.get('https://wing.coupang.com/delivery/management#?coupangSrl=&platformType=&from=2018-12-24&to=2019-01-22&deliveryStatus=ACCEPT&detailConditionKey=ORDER_ID&detailConditionValue=&vendoritempackgeName=&page=1&maxPage=1&cnt=1')

driver.find_element_by_name('id').send_keys('tg7616')
driver.find_element_by_name('pw').send_keys('p@ssw0rd6806!')
driver.find_element_by_xpath('//*[@id="btnLogin"]').click()

#element는 최초발견 1개만, elements는 전체발견함
order_No = driver.find_elements_by_css_selector('td:nth-child(2) > a.btn-order-simpleview')
bill_No = driver.find_elements_by_css_selector('td:nth-child(3)')
release_Day = driver.find_elements_by_css_selector('td:nth-child(4)')
product_Name = driver.find_elements_by_css_selector('td:nth-child(5)')
quantity = driver.find_elements_by_css_selector('td:nth-child(6)')
customer_Name = driver.find_elements_by_css_selector('td:nth-child(7)')
shipping = driver.find_elements_by_css_selector('td:nth-child(8)')
customer_Nomber = driver.find_elements_by_css_selector('td:nth-child(9)')
payment_Day = driver.find_elements_by_css_selector('td:nth-child(10)')
progress = driver.find_elements_by_css_selector('td:nth-child(11)')

print(order_No)
print(bill_No)
print(release_Day)
print(product_Name)
print(quantity)
print(customer_Name)
print(shipping)
print(customer_Nomber)
print(payment_Day)
print(progress)

#json
import json
from collections import OrderedDict

file_name = min(len(order_No), len(bill_No), len(release_Day), len(product_Name), len(customer_Name), len(quantity), len(shipping), len(customer_Nomber), len(payment_Day), len(progress))
parsing = OrderedDict()

for index in range(0, file_name):
    parsing['번호' + str(index + 1)] = {

        '주문번호': order_No[index].text,
        '운송장번호': bill_No[index].text,
        '출고예정일': release_Day[index].text,
        '등록상품명': product_Name[index].text,
        '구매수량': quantity[index].text,
        '구매자': customer_Name[index].text,
        '배송': shipping[index].text,
        '구매자전화번호': customer_Nomber[index].text,
        '결제일':payment_Day[index].text,
        '진행': progress[index].text
    }

with open('cupang.json', 'w', encoding="utf-8") as pa:
    json.dump(parsing, pa, ensure_ascii=False, indent="\t")

import pandas

pandas.read_json('cupang.json', encoding='utf8', orient='index').to_excel('cupang.xlsx', encoding='utf8')