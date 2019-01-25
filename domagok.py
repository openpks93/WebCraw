from selenium import webdriver
driver = webdriver.Chrome(r'C:\Program Files (x86)\Google\Chrome\chromedriver.exe')

driver.implicitly_wait(5)

driver.get('http://domeggook.com/main/mySell/sell/my_sellList.php?mode=WAITOK')


driver.find_element_by_name('id').send_keys('tg7616')
print("ok")
driver.find_element_by_name('pass').send_keys('aseman6806!')
print("ok")
driver.find_element_by_xpath('//*[@id="loginInput"]/input[3]').click()

#element는 최초발견 1개만, elements는 전체발견함
order_No = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(1)')
sale_Con = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(2) > font')
product_Name = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(3) > a')
customer_Name = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(4)')
quantity = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(5)')
payment = driver.find_elements_by_css_selector('#v5ContentsWrap > form > table:nth-child(4) > tbody > tr:nth-child(3) > td:nth-child(6) > div:nth-child(1)')


print(order_No)
print(sale_Con)
print(product_Name)
print(customer_Name)
print(quantity)
print(payment)

#json
import json
from collections import OrderedDict

file_name = min(len(order_No), len(sale_Con), len(product_Name), len(customer_Name), len(quantity), len(payment))
parsing = OrderedDict()

for index in range(0, file_name):
    parsing['no_' + str(index + 1)] = {

        '주문번호': order_No[index].text,
        '판매상태': sale_Con[index].text,
        '상품번호/상품제목': product_Name[index].text,
        '구매자': customer_Name[index].text,
        '수량': quantity[index].text,
        '결제금액': payment[index].text
    }

with open('domeggok.json', 'w', encoding="utf-8") as pa:
    json.dump(parsing, pa, ensure_ascii=False, indent="\t")

import pandas

pandas.read_json('domeggok.json', encoding='utf8', orient='index').to_excel('domeggok.xlsx', encoding='utf8')