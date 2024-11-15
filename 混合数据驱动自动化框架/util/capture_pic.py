from util.dir_util import create_date_hour_dir
from util.time_util import get_chinese_time
from config.var_config import  error_capture_pics_dir_path
import os
from selenium import  webdriver

def capture_pic(driver):
    try:
        dir_path = create_date_hour_dir(error_capture_pics_dir_path)
        pic_path = os.path.join(dir_path,get_chinese_time()+".png")
        driver.get_screenshot_as_file(pic_path)
        return pic_path
    except Exception as e:
        print("出现截图的异常：%s" %e)
        return ""

if __name__ == "__main__":
    driver = webdriver.Chrome(executable_path="e:\\chromedriver.exe")
    driver.get("https://www.sogou.com")
    pic_path = capture_pic(driver)
    print(pic_path)