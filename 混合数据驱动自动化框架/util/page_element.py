from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
def get_element(driver,locate_method,locate_exp):
    try:
        # 设置隐式等待时间为 10 秒
        if "xpath" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_xpath(locate_exp)
        elif "id" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_id(locate_exp)
        elif "tag_name" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_tag_name(locate_exp)
        elif "name" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_name(locate_exp)
        elif "partial_link" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_partial_link_text(locate_exp)
        elif "link_text" in str(locate_method):
            driver.implicitly_wait(6)
            element = driver.find_element_by_link_text(locate_exp)
        return element
    except NoSuchElementException:
        print("定位方法：%s  定位表达式:%s 无法定位到元素" % (locate_method, locate_exp))
        raise NoSuchElementException("定位方法：%s  定位表达式:%s 无法定位到元素" % (locate_method, locate_exp))
    except Exception as e:
        print(e)
        raise Exception("定位方法：%s  定位表达式:%s 进行定位时候出现异常" % (locate_method, locate_exp))

def get_elements(driver,locate_method,locate_exp):
    try:
        # 设置隐式等待时间为 10 秒
        if "xpath" in str(locate_method):
            driver.implicitly_wait(6)
            elements = driver.find_elements_by_xpath(locate_exp)
        elif "id" in str(locate_method):
            driver.implicitly_wait(6)
            elements = driver.find_elements_by_id(locate_exp)
        elif "tag_name" in str(locate_method):
            driver.implicitly_wait(6)
            elements = driver.find_elements_by_tag_name(locate_exp)
        elif "name" in str(locate_method):
            driver.implicitly_wait(6)
            elements = driver.find_elements_by_name(locate_exp)
        elif "partial_link" in str(locate_method):
            driver.implicitly_wait(6)
            elements= driver.find_elements_by_partial_link_text(locate_exp)
        elif "link_text" in str(locate_method):
            driver.implicitly_wait(6)
            elements = driver.find_elements_by_link_text(locate_exp)
        return elements
    except NoSuchElementException:
        print("定位方法：%s  定位表达式:%s 无法定位到元素" % (locate_method, locate_exp))
        raise NoSuchElementException("定位方法：%s  定位表达式:%s 无法定位到元素" % (locate_method, locate_exp))
    except Exception as e:
        print(e)
        raise Exception("定位方法：%s  定位表达式:%s 进行定位时候出现异常" % (locate_method, locate_exp))

if __name__ =="__main__":
    driver = webdriver.Chrome(executable_path="e:\\chromedriver.exe")
    driver.get("https://www.sogou.com")
    print(get_element(driver,"xpath","//input[@id='query']"))
    print(get_element(driver, "id","query"))
    print(get_element(driver, "name","query"))
    print(get_element(driver, "partial_link_text","输入法"))
    print(get_element(driver, "link_text","微信"))
    print(get_element(driver, "tag_name", "a"))
    #print(get_element(driver, "xpath", "//input[@id='query1']"))

    print(get_elements(driver,"xpath","//input[@id='query']"))
    print(get_elements(driver, "id","query"))
    print(get_elements(driver, "name","query"))
    print(get_elements(driver, "partial_link_text","输入法"))
    print(get_elements(driver, "link_text","微信"))
    print(get_elements(driver, "tag_name", "a"))
    driver.close()