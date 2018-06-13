# -*- coding:utf-8 -*-
from selenium import webdriver
import time,re,os
from bs4 import BeautifulSoup
import pandas as pd

path = os.getcwd().replace('\\','/')+'/'


def open_browser(url):
    driver = webdriver.Chrome('D:/python/docu/chromedriver.exe')
    driver.get(url)
    return driver


def log_in(driver):
    # 模拟登陆
    driver.find_element_by_xpath(
        ".//*[@id='web-content']/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/input"). \
        send_keys(username)
    driver.find_element_by_xpath(
        ".//*[@id='web-content']/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[3]/input"). \
        send_keys(password)
    driver.find_element_by_xpath(
        ".//*[@id='web-content']/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[5]").click()
    time.sleep(3)
    return driver


def search_company(driver, url1):
    driver.get(url1)
    content = driver.page_source.encode('utf-8')
    soup1 = BeautifulSoup(content, 'lxml')
    time.sleep(3)
    # company_name = soup1.find('a', class_="query_name sv-search-company f18 in-block vertical-middle").get_text()

    url2 = soup1.find('a', class_="query_name sv-search-company f18 in-block vertical-middle").attrs['href']
    driver.get(url2)
    # content2 = driver.page_source.encode('utf-8')
    # soup2 = BeautifulSoup(content2, 'lxml')
    return driver


def get_base_info(driver):
    base_table = {}
    base_table[u'名称'] = driver.find_element_by_xpath("//div[@class='company_header_width ie9Style position-rel']/div").text.split(u'我要认证')[0]
    base_info = driver.find_element_by_xpath("//div[@class='company_header_interior pl10 pt10 pb10 position-rel company-claim-header-bc mt15']")
    base_table[u'电话'] = base_info.text.split(u'电话：')[1].split(u'邮箱：')[0]
    base_table[u'邮箱'] = base_info.text.split(u'邮箱：')[1].split('\n')[0]
    base_table[u'网址'] = base_info.text.split(u'网址：')[1].split(u'地址')[0]
    base_table[u'地址'] = base_info.text.split(u'地址：')[1].split('\n')[0]
    abstract = driver.find_element_by_xpath("//div[@class='sec-c2 over-hide']/script")
    base_table[u'简介'] = driver.execute_script("return arguments[0].textContent", abstract).strip()
    tabs = driver.find_elements_by_tag_name('table')

    rows1 = tabs[0].find_elements_by_tag_name('tr')
    base_table[u'法人代表'] = rows1[1].find_elements_by_tag_name('td')[0].text.split('\n')[0]
    base_table[u'注册资本'] = rows1[1].find_elements_by_tag_name('td')[1].text.split('\n')[1]
    base_table[u'注册时间'] = rows1[1].find_elements_by_tag_name('td')[1].text.split('\n')[3]
    base_table[u'公司状态'] = rows1[1].find_elements_by_tag_name('td')[1].text.split('\n')[5]

    rows2 = tabs[1].find_elements_by_tag_name('tr')
    base_table[u'工商注册号'] = rows2[0].find_elements_by_tag_name('td')[1].text
    base_table[u'统一信用代码'] = rows2[1].find_elements_by_tag_name('td')[1].text
    base_table[u'纳税人识别号'] = rows2[2].find_elements_by_tag_name('td')[1].text
    base_table[u'营业期限'] = rows2[3].find_elements_by_tag_name('td')[1].text
    base_table[u'登记机关'] = rows2[4].find_elements_by_tag_name('td')[1].text

    base_table[u'组织机构代码'] = rows2[0].find_elements_by_tag_name('td')[3].text
    base_table[u'公司类型'] = rows2[1].find_elements_by_tag_name('td')[3].text
    base_table[u'行业'] = rows2[2].find_elements_by_tag_name('td')[3].text
    base_table[u'核准日期'] = rows2[3].find_elements_by_tag_name('td')[3].text
    base_table[u'英文名称'] = rows2[4].find_elements_by_tag_name('td')[3].text

    base_table[u'注册地址'] = rows2[5].find_elements_by_tag_name('td')[1].text.split(u'附近公司')[0]
    base_table[u'经营范围'] = rows2[6].find_elements_by_tag_name('td')[1].text

    return pd.DataFrame([base_table])


def get_staff_info(driver):
    staff_list = []
    staff_info = driver.find_elements_by_xpath("//div[@class='in-block f14 new-c5 pt9 pl10 overflow-width vertival-middle new-border-right']")
    for i in range(len(staff_info)):
        position = driver.find_elements_by_xpath("//div[@class='in-block f14 new-c5 pt9 pl10 overflow-width vertival-middle new-border-right']")[i].text
        person = driver.find_elements_by_xpath("//a[@class='overflow-width in-block vertival-middle pl15 mb4']")[i].text
        staff_list.append({u'职位': position, u'人员名称': person})
    staff_table = pd.DataFrame(staff_list, columns=[u'职位', u'人员名称'])
    return staff_table


def tryonclick(table):
    # 测试是否有翻页
    try:
        # 找到有翻页标记
        table.find_element_by_tag_name('ul')
        onclickflag = 1
    except Exception:
        print(u"没有翻页")
        onclickflag = 0
    return onclickflag


def change_page(table, df):
    PageCount = table.find_element_by_class_name('total').text
    PageCount = re.sub("\D", "", PageCount)  # 使用正则表达式取字符串中的数字 ；\D表示非数字的意思
    for i in range(int(PageCount) - 1):
        button = table.find_element_by_xpath(".//li[@class='pagination-next  ']/a")
        driver.execute_script("arguments[0].click();", button)
        time.sleep(3)
        df2 = get_table_info(table)
        df = df.append(df2)
    return df


def get_table_info(table):
    tab = table.find_element_by_tag_name('table')
    df = pd.read_html('<table>' + tab.get_attribute('innerHTML') + '</table>')
    if isinstance(df, list):
        df = df[0]
    if u'操作' in df.columns:
        del df[u'操作']
    return df


def scrapy(driver):
    tables = driver.find_elements_by_xpath("//div[contains(@id,'_container_')]")

    # 获取每个表格的名字
    c = '_container_'
    name = [0] * (len(tables) - 2)
    # 生成一个独一无二的十六位参数作为公司标记，一个公司对应一个，需要插入多个数据表
    id = keyword
    table_dict = {}
    for x in range(len(tables)-2):
        name[x] = tables[x].get_attribute('id')
        name[x] = name[x].replace(c, '')  # 可以用这个名称去匹配数据库
        # 判断是表格还是表单
        num = tables[x].find_elements_by_tag_name('table')

        # 基本信息表table有两个
        if len(num) > 1:
            table_dict[name[x]] = get_base_info(driver)

        elif name[x] in ['recruit', 'tmInfo']:
            pass

        #  单纯的表格
        elif len(num) == 1:
            df = get_table_info(tables[x])

            onclickflag = tryonclick(tables[x])
            # 判断此表格是否有翻页功能
            if onclickflag == 1:
                df = change_page(tables[x], df)

            table_dict[name[x]] = df

        # 表单样式
        elif name[x] == 'staff':
            table_dict[name[x]] = get_staff_info(driver)

        else:
            table_dict[name[x]] = pd.DataFrame()

    table_dict['websiteRecords'] = get_table_info(tables[len(tables)-2])

    return table_dict


def gen_excel(table_dict, keyword):
    with pd.ExcelWriter(path+keyword+'.xlsx') as writer:
        for sheet_name in table_dict:
            table_dict[sheet_name].to_excel(writer, sheet_name=sheet_name, index=None)


if __name__ == "__main__":
    url = 'https://www.tianyancha.com/login'
    url1 = 'http://www.tianyancha.com/search?key=%s&checkFrom=searchBox' % keyword

    driver = open_browser(url)
    driver = log_in(driver)
    driver = search_company(driver, url1)
    table_dict = scrapy(driver)
    gen_excel(table_dict, keyword)
