from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import xlwt

browser = webdriver.PhantomJS()
WAIT = WebDriverWait(browser,10)
browser.set_window_size(1400,900)

book = xlwt.Workbook(encoding='utf-8',style_compression=0)

sheet = book.add_sheet('蔡徐坤篮球',cell_overwrite_ok=True)

headerTitles = [u'名称',u'地址',u'描述',u'观看次数',u'弹幕数',u'发布时间']

for index in range(len(headerTitles)):
    sheet.write(0,index,headerTitles[index])

n = 1

def search():
    try:
        print('开始访问B站')
        browser.get('https://www.bilibili.com/')

        # 处理登录
        index = WAIT.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#primary_menu > ul > li.home > a")))
        index.click()

        input = WAIT.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#banner_link > div > div > form > input")))
        submit = WAIT.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="banner_link"]/div/div/form/button')))

        input.send_keys("蔡徐坤 篮球")
        submit.click()

        # 跳转到新的窗口
        print("跳转到新窗口")
        all_h = browser.window_handles
        browser.switch_to_window(all_h[1])

        get_source()
        total = WAIT.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                           "#server-search-app > div.contain > div.body-contain > div > div.page-wrap > div > ul > li.page-item.last > button")))
        return int(total.text)
    except TimeoutException:
        return search()

def next_page(page_num):
    try:
        print('获取下一页数据')
        next_btn = WAIT.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                          '#server-search-app > div.contain > div.body-contain > div > div.page-wrap > div > ul > li.page-item.next > button')))
        next_btn.click()
        WAIT.until(EC.text_to_be_present_in_element((By.CSS_SELECTOR,
                                                     '#server-search-app > div.contain > div.body-contain > div > div.page-wrap > div > ul > li.page-item.active > button'),
                                                    str(page_num)))
        get_source()
    except TimeoutException:
        browser.refresh()
        return next_page(page_num)

def save_to_excel(soup):
    list = soup.find(class_='all-contain').find_all(class_='info')
    for item in list:
        item_title = item.find('a').get('title')
        item_link = item.find('a').get('href')
        item_dec = item.find(class_='des hide').text
        item_view = item.find(class_='so-icon watch-num').text
        item_biubiu = item.find(class_='so-icon hide').text
        item_date = item.find(class_='so-icon time').text

        print('爬取:' + item_title)

        global n

        sheet.write(n,0,item_title)
        sheet.write(n,1,item_link)
        sheet.write(n,2,item_dec)
        sheet.write(n,3,item_view)
        sheet.write(n,4,item_biubiu)
        sheet.write(n,5,item_date)

        n = n + 1

def get_source():
    WAIT.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#server-search-app > div.contain > div.body-contain > div > div.result-wrap.clearfix')))
    html = browser.page_source
    soup = BeautifulSoup(html, 'lxml')
    save_to_excel(soup)

def main():
    try:
        total = search()
        print(total)

        for i in range(2,int(total + 1)):
            next_page(i)
    
    finally:
        browser.close()

if __name__ == '__main__':
    main()
    book.save(u'蔡徐坤篮球.xlsx')