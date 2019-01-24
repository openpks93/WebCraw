from selenium import webdriver
print("set url : " + 'http://ipss.interpark.com/ipss/ipssmainscr.do?_method=initial&_style=ipssPro&wid1=wgnb&wid2=wel_login&wid3=seller')
# Chrome Headless 전용 옵션, Firefox 사용시 모두 주석
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# Chrome 드라이버 생성(둘 중 하나만 켤것)
chrome_driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
chrome_driver.implicitly_wait(5)
chrome_driver.get('http://ipss.interpark.com/ipss/ipssmainscr.do?_method=initial&_style=ipssPro&wid1=wgnb&wid2=wel_login&wid3=seller')

# 아이디/비밀번호를 입력해준다.
chrome_driver.find_element_by_id('memId').send_keys('tg7616')
chrome_driver.find_element_by_id('pwd').send_keys('aseman6806!')
chrome_driver.find_element_by_xpath('//*[@id="tab1"]/div[1]/div[2]/button').click()

print("webdriver's ready")
# 크롤링할 사이트 호출
print("getting url site")
print("parsing data")

Item_number = chrome_driver.find_elements_by_css_selector("div.jqx-clear > div.jqx-grid-content > div > div > div.jqx-grid-cell > div > a")
product_name = chrome_driver.find_elements_by_css_selector("div.jqx-clear > div.jqx-grid-content > div > div > div.jqx-grid-cell > div > a")

print(Item_number)

# # Json
# import json
# from collections import OrderedDict
# LENGTH = min(len(Item_number), len(product_name))
# products = OrderedDict()
# print("processing data to json")
#
# for idx in range(0, LENGTH):
#     products['no_' + str(idx + 1)] = {
#         'Item_Order_Number': Item_number[idx].text,
#         'Order_Number': product_name[idx].text
#     }
#
#
# ###################### json 저장
# with open('interpark.json', 'w', encoding='utf-8') as f:
#     json.dump(products, f, ensure_ascii=False, indent="\t")
#
# ####################### xlsx 형식으로 만들기
# import pandas
# pandas.read_json('interpark.json',  encoding='UTF8').to_excel('interpark.xlsx', encoding='UTF8')
#
# ####################### 브라우저 종료, 웹 드라이버 종료
# print("This job is finished and close the web browser")
# chrome_driver.quit()

