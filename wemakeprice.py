from selenium import webdriver

print("set url : " + 'http://seller.nhmarket.kr/mallinmall/order/order_list.asp?menu=delivery')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(20)
chrome_driver.get('http://biz.wemakeprice.com/dealer/settlement/sales_list')
# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('login_uid').send_keys('tg7616')
chrome_driver.find_element_by_id('login_upw_biz').send_keys('p@ssw0rd6806!')
chrome_driver.find_element_by_xpath('//*[@id="loginConfirmBtn_biz"]').click()
#조회 클릭
chrome_driver.find_element_by_xpath('/html/body/div[11]/div[1]/div[1]/button').click()
chrome_driver.find_element_by_xpath('/html/body/div[11]/div[1]/h1/a/img').click()
chrome_driver.find_element_by_xpath('//*[@id="partner_popup"]/div[2]/a').click()
chrome_driver.find_element_by_xpath('//*[@id="agree"]/div[2]/a').click()
chrome_driver.find_element_by_xpath('/html/body/div[11]/div[1]/ul/li[3]/a').click()
chrome_driver.find_element_by_xpath('/html/body/div[11]/div[1]/ul/li[3]/div/ul/li[1]/a').click()
chrome_driver.find_element_by_xpath('//*[@id="group_month"]/label').click()
chrome_driver.find_element_by_xpath('//*[@id="btn_search"]').click()
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

# 크롤링 element
#orderDay = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(2)')
deal_id = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(3)')
deal_name = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(4)')
settlement_Cycle = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(5)')
payment_method = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(6)')
period_of_sale = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(7)')
logistics_division = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(8)')
shipping_Quantity = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(9)')
refund_Quantity = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(10)')
total_shipped_amount = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(11)')
refund_amount = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(12)')
sales_agent_fee = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(13)')
wemake_Item_coupon = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(14)')
owner_Item_coupon= chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(15)')
shipping_fee = chrome_driver.find_elements_by_css_selector('#search_result_table > tbody > tr > td:nth-child(16)')

# Json
import json
from collections import OrderedDict
LENGTH = min(len(deal_id), len(deal_name), len(settlement_Cycle), len(payment_method), len(period_of_sale), len(logistics_division), len(shipping_Quantity), len(refund_Quantity), len(total_shipped_amount), len(refund_amount), len(sales_agent_fee), len(wemake_Item_coupon), len(owner_Item_coupon), len(shipping_fee))
products = OrderedDict()
print("processing data to json")

for idx in range(0, LENGTH):
    products['no_' + str(idx + 1)] = {
        'deal_id': deal_id[idx].text,
        'deal_name': deal_name[idx].text,
        'settlement_Cycle': settlement_Cycle[idx].text,
        'payment_method': payment_method[idx].text,
        'period_of_sale': period_of_sale[idx].text,
        'logistics_division': logistics_division[idx].text,
        'shipping_Quantity': shipping_Quantity[idx].text,
        'refund_Quantity': refund_Quantity[idx].text,
        'total_shipped_amount': total_shipped_amount[idx].text,
        'refund_amount': refund_amount[idx].text,
        'sales_agent_fee': sales_agent_fee[idx].text,
        'wemake_Item_coupon': wemake_Item_coupon[idx].text,
        'owner_Item_coupon': owner_Item_coupon[idx].text,
        'shipping_fee': shipping_fee[idx].text
    }

# json 저장
with open('wemakeprice.json', 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent="\t")

# xlsx 형식으로 만들기
import pandas
pandas.read_json('wemakeprice.json',  encoding='UTF8', orient='index').to_excel('wemakeprice.xlsx', encoding='UTF8')

# 브라우저 종료, 웹 드라이버 종료
print("This job is finished and close the web browser")
chrome_driver.quit()

