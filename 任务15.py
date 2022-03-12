import os
import pickle
import re
import openpyxl
import requests
from selenium import webdriver
import time
import warnings
from lxml import etree
from pyquery import PyQuery as pq
warnings.filterwarnings("ignore", category=DeprecationWarning)


class Postman:

    def __init__(self):
        self.url = 'https://www.bypms.cn/'
        self.get_data()
        self.driver = webdriver.Chrome(r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver")
        self.login()
        self.run()

    def getCookies(self):
        self.driver.get(self.url)
        time.sleep(1)
        phone = self.driver.find_element_by_xpath('//*[@id="frmLogin"]/div[1]/span[2]')
        phone.click()
        username = self.driver.find_element_by_id('mobile')
        username.click()
        username.clear()
        username.send_keys('17898834046')
        password = self.driver.find_element_by_id('password')
        password.click()
        password.clear()
        password.send_keys('123456')
        login = self.driver.find_element_by_class_name('bgR')
        login.click()
        time.sleep(1)
        my_Cookies = self.driver.get_cookies()
        self.driver.quit()
        cookies = {}
        for item in my_Cookies:
            cookies[item['name']] = item['value']
        outputPath = open('Cookies.pickle', 'wb')
        pickle.dump(cookies, outputPath)
        outputPath.close()
        return cookies

    def read_cookies(self):
        if os.path.exists('Cookies.pickle'):
            readPath = open('Cookies.pickle', 'rb')
            Cookies = pickle.load(readPath)
        else:
            Cookies = self.getCookies()
        return Cookies

    def login(self):
        time.sleep(1)
        self.driver.maximize_window()
        Cookies = self.read_cookies()
        self.driver.get(self.url)
        for cookie in Cookies:
            self.driver.add_cookie({
                "domain": ".bypms.cn",
                "name": cookie,
                "value": Cookies[cookie],
                "path": '/',
                "expires": None
            })
        self.driver.refresh()
        time.sleep(1)
        self.driver.switch_to.window(self.driver.window_handles[0])
        self.driver.find_element_by_xpath('//*[@id="frmLogin"]/div[1]/a').click()
        time.sleep(1)
        self.driver.switch_to.window(self.driver.window_handles[0])
        try:
            self.driver.find_element_by_xpath('//*[@id="_license_notice"]/section/div[3]/span[1]').click()
        except:
            pass

    def get_data(self):
        url = 'https://www.bypms.cn/console/room/'
        headers = {
            'cookie': '_dhsid_dh_sr_v3=7ml9co8spy0zna53; JSESSIONID=BA50D2415709E9B50942E9F9EC7EB441; Hm_lvt_b03b3baab5c864cd81ddcd129ce51234=1645454756,1646406874,1646843811,1646925187; Hm_lpvt_b03b3baab5c864cd81ddcd129ce51234=1646925192',
            'Host': 'www.bypms.cn',
            'Referer': 'https://www.bypms.cn/console/channel/unit/?roomIds=no&keyword=%E8%97%93%E5%9B%AD%E6%95%B4%E6%A0%8B10%E9%97%B4',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.36',
        }
        response = requests.get(url, headers=headers)
        doc = pq(response.text)
        trs = doc('.itable tr')
        with open('data.txt', 'w', encoding='utf-8') as f:
            for tr in trs.items():
                room = tr('.r_no').text()
                f.write(room)
                f.write('\n')

    def run(self):
        url = 'https://www.bypms.cn/console/channel/unit/?roomIds=no'
        self.driver.get(url)
        time.sleep(1)
        self.driver.switch_to.window(self.driver.window_handles[0])
        with open('data.txt', 'r', encoding='utf-8') as f:
            rooms = f.read()
            rooms = rooms.split('\n')
            print(rooms)
        for room in rooms:
            if room:
                search = self.driver.find_element_by_xpath('//*[@placeholder="关键词"]')
                search.click()
                search.clear()
                search.send_keys(room)
                btn = self.driver.find_element_by_xpath('//*[@value="查找"]')
                btn.click()
                try:
                    self.driver.find_element_by_xpath('//*[@id="hd_lstUnit"]/table/thead/tr/td[1]').click()
                    time.sleep(1)
                    self.driver.switch_to.window(self.driver.window_handles[0])
                    self.driver.find_element_by_xpath('//*[@id="bar_lstUnit"]/span').click()
                    time.sleep(1)
                    self.driver.switch_to.window(self.driver.window_handles[0])
                    self.driver.switch_to.frame('ifrm__select_rooms')
                    doc = pq(self.driver.page_source)
                    labels = doc('#vchecker label').items()
                    for label in labels:
                        if label.attr('title') == room:
                            self.driver.find_element_by_xpath(f'//*[@title="{room}"]').click()
                            time.sleep(1)
                            self.driver.find_element_by_id('btnSubmit').click()
                            self.driver.get(url)
                            time.sleep(1)
                            self.driver.switch_to.window(self.driver.window_handles[0])
                except:
                    pass
                finally:
                    time.sleep(2)


if __name__ == '__main__':
    Postman()
