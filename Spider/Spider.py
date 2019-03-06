# -*- coding: utf-8 -*-
import logging
import time
import json
import sys
import argparse
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import requests
import random
import xlsxwriter


class SpiderDynamic(object):
    def __init__(self):
        opt = webdriver.ChromeOptions()
        # opt.add_argument('--headless')
        # opt.add_argument('blink-settings=imagesEnabled=false')
        self.browser = webdriver.Chrome(chrome_options=opt)

    def visit_site(self, url):
        """
        访问页面
        :param url:
        :return:
        """
        try:
            self.browser.get(url)
        except:
            logging.error("Can not visit this site %s" % url)
            return False
        return True

    def get_text_by_xpath(self, xpath):
        """
        获取路径下的文本信息
        :param xpath:
        :return:
        """
        try:
            data = self.browser.find_element_by_xpath(xpath=xpath).text
        except:
            logging.error("No such elem %s to get text" % xpath)
            return None
        return data

    def get_attribute_by_xpath(self, xpath, key):
        """
        获取路径下的文本信息
        :param xpath:
        :return:
        """
        try:
            data = self.browser.find_element_by_xpath(xpath=xpath).get_attribute(key)
        except:
            logging.error("No such elem %s to get text" % xpath)
            return None
        return data

    def is_exist_xpath(self, xpath):
        """
        检查当前的xpath是否存在
        :param xpath:
        :return:
        """
        try:
            self.browser.find_element_by_xpath(xpath=xpath)
        except:
            return False
        return True

    def set_select_visible_text_by_xpath(self, xpath, value):
        """
        select类的下拉框选择visible text
        :param xpath:
        :param value:
        :return:
        """
        try:
            Select(self.browser.find_element_by_xpath(xpath=xpath)).select_by_visible_text(value)
        except:
            logging.error("Failed to set value to select")
            return False
        return True

    def set_text_by_xpath(self, xpath, value):
        """
        设置value到xpath
        :param xpath:
        :param value:
        :return:
        """
        try:
            self.browser.find_element_by_xpath(xpath).send_keys(value)
        except:
            logging.error("Failed to set value")
            return False
        return True

    def click_by_xpath(self, xpath):
        """
        点击某个路径
        :param xpath:
        :return:
        """
        try:
            self.browser.find_element_by_xpath(xpath=xpath).click()
        except:
            logging.error("No such elem %s to be clicked" % xpath)
            return False
        return True

    def close_current_tab(self):
        """
        关闭当前标签页
        :return:
        """
        self.browser.close()

    def switch_to_nearest_page(self):
        """
        焦点切换到最近打开的页面
        :return:
        """
        try:
            self.browser.switch_to.window(self.browser.window_handles[-1])
        except:
            logging.error("Failed to switch to the nearest page")
            return False
        return True

    def switch_to_first_page(self):
        """
        焦点切换到第一次打开的页面
        :return:
        """
        try:
            self.browser.switch_to.window(self.browser.window_handles[0])
        except:
            logging.error("Failed to switch to the first page")
            return False
        return True

    def disable_readonly(self, id):
        """

        :param xpath:
        :return:
        """
        js = 'document.getElementById(%s).removeAttribute("readonly");' % id
        self.browser.execute_script(js)

    def set_value_by_js(self, id, value):
        """
        通过js代码写入值
        :param id:
        :param value:
        :return:
        """
        js = 'document.getElementById(%s).value="%s"' % (id, value)
        self.browser.execute_script(js)

    def __del__(self):
        """
        close all
        :return:
        """
        self.browser.close()
        self.browser.quit()


class BidingSpiderDynamic(SpiderDynamic):
    def __init__(self, config_file="config.json", keywords_file="keywords.dic", fix_date=""):
        super(BidingSpiderDynamic, self).__init__()
        with open(config_file, "r") as fp:
            self.__sites = json.load(fp)

        self.__keywords = set()
        with open(keywords_file, "r") as fp:
            for line in fp.readlines():
                self.__keywords.add(line.strip("\r\n"))

        if fix_date == "":
            self.__current_date_ts = int(time.strftime("%Y%m%d", time.localtime(time.time())))
            self.__current_date_str = time.strftime("%Y-%m-%d", time.localtime(time.time()))
        else:
            self.__current_date_ts = int(fix_date.replace("-", ""))
            self.__current_date_str = fix_date

        self.__user_agent_list = [
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
            "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
            "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5",
            "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
            "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
            "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24"
        ]

    def __parse_static_page(self, url, params):
        """
        解析静态页面
        :param url:
        :param params:
        :return:
        """
        headers = {
            'User-Agent': random.choice(self.__user_agent_list),
            'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            'Accept-Encoding': 'gzip, deflate, sdch',
        }
        try:
            r = requests.get(url=url, headers=headers, params=params)
            r.encoding = "utf-8"
            bs = BeautifulSoup(r.text, "html.parser")
        except:
            logging.error("Failed to get %s" % url)
            return None
        return bs

    def __cmcc_page(self, name, url):
        results = []
        next_page_xpath = "//*[@id=\"pageid2\"]/table/tbody/tr/td[4]/a"
        if self.visit_site(url) is False:
            logging.error("Faile to visit %s" % url)
            return results
        time.sleep(3)
        self.set_text_by_xpath("//*[@id=\"startDate\"]", self.__current_date_str)
        self.set_text_by_xpath("//*[@id=\"endDate\"]", self.__current_date_str)
        self.click_by_xpath("//*[@id=\"search\"]")
        time.sleep(3)
        flag = 0
        while flag == 0 or self.is_exist_xpath(next_page_xpath):
            idx = 3
            flag = 1
            title_xpath = "//*[@id=\"searchResult\"]/table/tbody/tr[%d]/td[3]/a" % idx
            while self.is_exist_xpath(title_xpath):
                # 点击标题页
                if self.click_by_xpath(title_xpath) is False:
                    logging.warning("Failed to click title %s" % title_xpath)
                    break
                else:
                    time.sleep(3)
                    self.switch_to_nearest_page()
                    title = self.get_text_by_xpath("//*[@id=\"container\"]/div[1]/table/tbody/tr[2]/td/h1")
                    page_url = self.browser.current_url
                    self.close_current_tab()
                    self.switch_to_first_page()
                    results.append({"name": name, "title": title, "url": page_url})

                    idx = idx + 1
                    title_xpath = "//*[@id=\"searchResult\"]/table/tbody/tr[%d]/td[3]/a" % idx
            if self.is_exist_xpath(next_page_xpath) is False:
                break
            self.click_by_xpath(next_page_xpath)
            time.sleep(3)

        return results

    def __chinatelecom_page(self, name, url):
        """
        1、访问页面
        1、填入需要查询的词语
        2、点击查询
        3、获取表单的结果
        //*[@id="two_pages_all"]/div[1]/div[2]/table[3]/tbody/tr[2]/td[3]/a
        //*[@id="two_pages_all"]/div[1]/div[2]/table[3]/tbody/tr[3]/td[3]/a
        //*[@id="two_pages_all"]/div[1]/div[2]/table[3]/tbody/tr[4]/td[3]/a
        :param name:
        :return:
        """
        value_xpath = '//*[@id="docTitle"]'
        search_xpath = '//*[@id="two_pages_all"]/div[1]/div[2]/table[1]/tbody/tr[4]/td/input[1]'
        start_time = self.__current_date_str + " 00:00:00"
        end_time = self.__current_date_str + " 23:55:45"
        start_ts = int(time.mktime(time.strptime(start_time, "%Y-%m-%d %H:%M:%S")))
        end_ts = int(time.mktime(time.strptime(end_time, "%Y-%m-%d %H:%M:%S")))
        self.visit_site(url)
        time.sleep(3)
        results = []
        next_page_xpath = "//*[@id=\"two_pages_all\"]/div[1]/div[2]/table[4]/tbody/tr/td[10]/a[1]"
        flag = 0
        while flag == 0 or self.is_exist_xpath(next_page_xpath):
            idx = 2
            flag = 1
            ts_xpath = "//*[@id=\"two_pages_all\"]/div[1]/div[2]/table[3]/tbody/tr[%d]/td[6]" % idx
            title_xpath = "//*[@id=\"two_pages_all\"]/div[1]/div[2]/table[3]/tbody/tr[%d]/td[3]/a" % idx
            while self.is_exist_xpath(ts_xpath):
                title = self.get_text_by_xpath(title_xpath)
                date_str = self.get_text_by_xpath(ts_xpath).strip()
                ts = int(time.mktime(time.strptime(date_str, "%Y-%m-%d %H:%M:%S")))
                if ts < start_ts or ts > end_ts:
                    logging.info("Out of date")
                    break

                # 点击标题页
                self.click_by_xpath(title_xpath)
                time.sleep(3)
                self.switch_to_nearest_page()
                page_url = self.browser.current_url
                self.close_current_tab()
                self.switch_to_first_page()
                results.append({"name": name, "title": title, "url": page_url})

                idx = idx + 1
                ts_xpath = "//*[@id=\"two_pages_all\"]/div[1]/div[2]/table[3]/tbody/tr[%d]/td[6]" % idx
                title_xpath = "//*[@id=\"two_pages_all\"]/div[1]/div[2]/table[3]/tbody/tr[%d]/td[3]/a" % idx
            if self.is_exist_xpath(next_page_xpath) is False:
                break
            self.click_by_xpath(next_page_xpath)
            time.sleep(3)
        return results

    def __zbytb_page(self, name, url):
        """
        1、访问网页
        2、while 日期是当天的
            while 当前页面的xpath存在
                搜集信息 /html/body/div[5]/div[1]/table/tbody/tr[2]/td[2]/a
                    /html/body/div[5]/div[1]/table/tbody/tr[3]/td[2]/a
                    /html/body/div[5]/div[1]/table/tbody/tr[4]/td[2]/a
                url 路径下的herf参数
                如果日期不是当天的，break
        :param name:
        :return:
        """
        results = []
        params = {"kw": "", "catid": "101"}
        for words in self.__keywords:
            params["keywords"] = words
            bs_page = self.__parse_static_page(url, params)
            if u"当前访问疑似黑客攻击，已被网站管理员设置为拦截" in bs_page.text:
                logging.warning("Anti-Spider")
                break
            check_flag = bs_page.select(".zblist_table")[0].find_all("tr")
            if len(check_flag) == 1:
                logging.warning("No record url=%s, keyword=%s" % (url, words))
                continue
            data_bs = bs_page.select(".zblist_table")[0].find_all("tr")
            for idx, elem in enumerate(data_bs):
                if idx == 0:
                    continue
                fileds = elem.find_all("td")
                title = fileds[1].text.encode("utf-8").strip("\r\n")
                page_url = fileds[1].a.attrs['href']
                time_data = int(fileds[-1].text.encode("utf-8").strip("\r\n").replace("-", ""))
                if time_data < self.__current_date_ts:
                    logging.info("Out of date")
                    break
                results.append({"name": name.decode("utf-8"), "title": title.decode("utf-8"), "url": page_url.decode("utf-8")})

        return results

    def __chinabidding_page(self, name, url):
        """
        :param name:
        :return:
        """
        site = "https://www.chinabidding.cn/"
        results = []
        params = {"keywords": "", "search_type": "TITLE"}
        for words in self.__keywords:
            params["keywords"] = words
            bs_page = self.__parse_static_page(url, params)
            check_flag = bs_page.select(".lieb")[0].table.tbody.tr.td.table.tbody.tr.text
            if u"一共搜索到 0 条记录" in check_flag:
                logging.warning("No record url=%s, keyword=%s" % (url, words))
                continue
            data_bs = bs_page.select(".lieb")[0].table.tbody.tr.td.table.tbody.find_all("tr")
            for elem in data_bs:
                title = elem.td.text.encode("utf-8").strip("\r\n")
                page_url = site + elem.td.a.attrs['href']
                time_data = int(elem.find_all("td")[-1].text.encode("utf-8").strip("\r\n").replace("-", ""))
                if time_data < self.__current_date_ts:
                    logging.info("Out of date")
                    break
                results.append({"name": name.decode("utf-8"), "title": title.decode("utf-8"), "url": page_url.decode("utf-8")})

        return results

    def __bidcenter_page(self, name, url):
        textbox_xpath = "/html/body/div[5]/div[1]/form/input[1]"
        search_btn_xpath = "/html/body/div[5]/div[1]/form/input[4]"
        catgray_xpath = "//*[@id=\"catid_1\"]"
        results = []
        logging.info("Visit bidcenter site")
        for words in self.__keywords:
            self.visit_site(url=url + "?keywords=%s&tag=1" % words)

            ts = int(time.strftime("%Y%m%d", time.localtime(time.time()))) - 1
            idx = 1
            time_str_xpath = "//*[@id=\"jq_project_list\"]/tbody/tr[%d]/td[7]" % idx
            title_xpath = "//*[@id=\"jq_project_list\"]/tbody/tr[%d]/td[2]/a" % idx

            while ts >= self.__current_date_ts and self.is_exist_xpath(time_str_xpath) is True:
                time_str_xpath = "//*[@id=\"jq_project_list\"]/tbody/tr[%d]/td[7]" % idx
                title_xpath = "//*[@id=\"jq_project_list\"]/tbody/tr[%d]/td[2]/a" % idx
                data = self.get_text_by_xpath(time_str_xpath)
                if data is None:
                    logging.error("Failed to get timestr")
                    return results
                ts = int(data.strip().replace("-", ""))
                if ts < self.__current_date_ts:
                    logging.info("Out of date")
                    break
                title = self.get_text_by_xpath(title_xpath)
                page_url = self.get_attribute_by_xpath(title_xpath, "href")
                results.append({"name": name, "title": title, "url": page_url})
                idx = idx + 1
            time.sleep(5)
        return results

    def run(self, output="data"):
        """
        爬取所有的招标页面动态页面
        :return:
        """
        book = xlsxwriter.Workbook(output + "-" + self.__current_date_str + ".xlsx")
        for elem in self.__sites:
            name = elem["name"].encode("utf-8")
            id = elem["id"].encode("utf-8")
            url = elem["url"].encode("utf-8")
            data = []
            logging.info("Run %s in %s" % (id, name))

            if id == "b2b.10086.cn":
                data = self.__cmcc_page(name, url)

            if id in ["caigou.chinatelecom.com.cn.jt", "caigou.chinatelecom.com.cn.njt"]:
                data = self.__chinatelecom_page(name, url)

            if id == "www.bidcenter.com.cn":
                data = self.__bidcenter_page(name, url)

            if id == "www.zbytb.com":
                data = self.__zbytb_page(name, url)

            if id == "www.chinabidding.cn":
                data = self.__chinabidding_page(name, url)

            if len(data) == 0:
                continue
            sheet = book.add_worksheet(name.decode("utf-8"))
            sheet.write(0, 0, "来源".decode("utf-8"))
            sheet.write(0, 1, "标题".decode("utf-8"))
            sheet.write(0, 2, "地址".decode("utf-8"))
            for idx,elem in enumerate(data):
                if type(elem["name"]) == str:
                    elem["name"] = elem["name"].decode("utf-8")
                if type(elem["title"]) == str:
                    elem["title"] = elem["title"].decode("utf-8")
                if type(elem["url"]) == str:
                    elem["url"] = elem["url"].decode("utf-8")
                sheet.write(idx + 1, 0, elem["name"])
                sheet.write(idx + 1, 1, elem["title"])
                sheet.write(idx + 1, 2, elem["url"])
        book.close()

        return None

    def __del__(self):
        pass


def basic_apply(arg_title, params):
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
        stream=sys.stderr
    )
    parser = argparse.ArgumentParser(description=arg_title)
    for elem in params:
        parser.add_argument(elem["param"], type=elem["type"], help=elem["help"])

    return parser.parse_args()


def main():
    params = [
        {"param": "sites", "type": type(""), "help": "Sites config"},
        {"param": "keywords", "type": type(""), "help": "Keywords dict"},
        {"param": "output", "type": type(""), "help": "output excel file"},
        {"param": "date_str", "type": type(""), "help": "Date [2019-03-05]"},
    ]
    args = basic_apply("Spider for biding sites", params)
    while True:
        bsd = BidingSpiderDynamic(config_file=args.sites, keywords_file=args.keywords)#, fix_date=args.date_str)
        bsd.run(output=args.output)
        del bsd
        time.sleep(24 * 3600)


if __name__ == '__main__':
    main()
