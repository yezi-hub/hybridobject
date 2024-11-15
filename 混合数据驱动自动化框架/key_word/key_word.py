import time

from selenium import webdriver
from util.page_element import get_element
from selenium.common.exceptions import NoSuchElementException
import traceback


def open_browser(browser_type):
    global driver
    if "edge" in browser_type.lower():
        driver = webdriver.Edge("e:\\edgedriver.exe")
    elif "chrome" in browser_type.lower():
        driver = webdriver.Chrome("e:\\chromedriver.exe")
    else:
        options = webdriver.FirefoxOptions()
        # 指定firefox.exe所在的绝对路径，不设定这个路径，可能无法启动
        options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"

        # 没有发生异常，表示在页面中找到了该元素，返回True
        # driver = webdriver.Chrome(executable_path="e:\\chromedriver.exe")
        driver = webdriver.Firefox(executable_path=r'e:\geckodriver.exe', options=options)
    return driver

def get(url):
    global driver
    print("-----------",driver)
    try:
        driver.get(url)
    except Exception as e:
        exception_info = traceback.format_exc()
        print("浏览器访问 url：%s 出现异常，异常信息：%s" %(url,exception_info))
        raise Exception("浏览器访问 url：%s 出现异常，异常信息：%s" %(url,exception_info))

def switch_to_iframe(locate_method,locate_exp):
    global driver
    try:
        element = get_element(driver,locate_method,locate_exp)
        driver.switch_to.frame(element)
    except NoSuchElementException as e:
        print("定位方式:%s 定位表达式：%s 没有定位到页面元素" %(locate_method,locate_exp))
        raise Exception("定位方式:%s 定位表达式：%s 没有定位到页面元素" %(locate_method,locate_exp))
    except Exception as e:
        print("定位方式:%s 定位表达式：%s 的iframe在切换时候出现异常" %(locate_method,locate_exp))
        exception_info = traceback.format_exc()
        raise Exception("定位方式:%s 定位表达式：%s 的iframe在切换时候出现异常,异常信息：\n" %(locate_method,locate_exp)+exception_info )

def switch_out_iframe():
    global driver
    try:
        driver.switch_to.default_content()
    except Exception as e:
        print("切出iframe的时候出现异常！")
        raise Exception("切出iframe的时候出现异常！")
def input(locate_method, locate_exp,content):
    global driver
    try:
        element = get_element(driver,locate_method, locate_exp)
        element.send_keys(content)
    except NoSuchElementException as e:
        print("定位方式:%s 定位表达式：%s 没有定位到页面元素" % (locate_method, locate_exp))
        raise Exception("定位方式:%s 定位表达式：%s 没有定位到页面元素" % (locate_method, locate_exp))
    except Exception as e:
        exception_info = traceback.format_exc()
        print("定位方式:%s 定位表达式：%s 的元素在输入内容 %s 时候出现异常，异常信息：%s" %(locate_method, locate_exp,content,exception_info))
        raise Exception("定位方式:%s 定位表达式：%s 的元素在输入内容 %s 时候出现异常，异常信息：%s" %(locate_method, locate_exp,content,exception_info))

def click(locate_method, locate_exp):
    global driver
    try:
        element = get_element(driver, locate_method, locate_exp)
        element.click()
    except NoSuchElementException as e:
        print("定位方式:%s 定位表达式：%s 没有定位到页面元素" %(locate_method, locate_exp))
        raise Exception("定位方式:%s 定位表达式：%s  没有定位到页面元素" %(locate_method, locate_exp))
    except Exception as e:
        exception_info = traceback.format_exc()
        print("定位方式:%s 定位表达式：%s 的元素在点击操作时候出现异常，异常信息：%s" % (locate_method, locate_exp,exception_info))
        raise Exception("定位方式:%s 定位表达式：%s 的元素在点击操作时候出现异常，异常信息：%s" % (locate_method, locate_exp,exception_info))

def click_check_box(locate_method, locate_exp,content):
    global driver
    if content == "是":
        star = True
    else:
        star = False
    try:
        element = get_element(driver, locate_method, locate_exp)
        if star and not element.is_selected():
            element.click()
    except NoSuchElementException as e:
        print("定位方式:%s 定位表达式：%s 没有定位到页面元素" %(locate_method, locate_exp))
        raise Exception("定位方式:%s 定位表达式：%s  没有定位到页面元素" %(locate_method, locate_exp))
    except Exception as e:
        exception_info = traceback.format_exc()
        print("定位方式:%s 定位表达式：%s 的元素在点击操作时候出现异常，异常信息：%s" % (locate_method, locate_exp,exception_info))
        raise Exception("定位方式:%s 定位表达式：%s 的元素在点击操作时候出现异常，异常信息：%s" % (locate_method, locate_exp,exception_info))



def sleep(seconds):
    seconds = int(seconds)
    time.sleep(seconds)

def assert_word(word):
    global  driver
    try:
        assert word in driver.page_source,"断言异常：断言词%s 没有在源码中被发现！" %word
    except AssertionError  as e:
        print("断言异常：断言词%s 没有在源码中被发现！" %word)
        raise e
def quit():
    global driver
    driver.quit()

if __name__ == "__main__":
    #open_browser("edge")
    driver = open_browser("chrome")
    #open_browser("firefox")
    get("https://mail.126.com")
    switch_to_iframe("xpath",'//div[@id="loginDiv"]/iframe')
    input("xpath",'//input[@name="email"]',"testman1980")
    input("xpath",'//input[@name="password"]',"wulaoshi1978")
    click("id","dologin")
    switch_out_iframe()
    sleep(5)
    click("xpath",'//span[.= "写 信"]')
    input("xpath",'//input[@tabindex="1" and @aria-label="收件人地址输入框，请输入邮件地址，多人时地址请以分号隔开"]',"testman1980@126.com")
    input("xpath",'//input[contains(@id,"_subjectInput")]',"今天天气不错")
    input("xpath",'//input[@type="file"]',"e:\\a.txt")
    sleep(3)
    switch_to_iframe("xpath",'//iframe[@tabindex="1"]')
    input("xpath",'//p',"下午没有课！")
    switch_out_iframe()
    click("xpath","//span[text()='发送']")
    sleep(3)
    assert_word("邮件发送成功")
    quit()


