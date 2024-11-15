import os
from util.time_util import *

def create_date_dir(dir_path):
    dir_path = os.path.join(dir_path,get_chinese_date())
    if not os.path.exists(dir_path):
        try:
            os.mkdir(dir_path)
            return dir_path
        except Exception as e:
            print("创建目录 %s 时候，出现异常： %s" %(dir_path,e))
    else:
        return dir_path

def create_date_hour_dir(dir_path):
    dir_path = create_date_dir(dir_path)
    date_hour_dir_path = os.path.join(dir_path,get_chinese_hour())
    if not os.path.exists(date_hour_dir_path):
        try:
            os.mkdir(date_hour_dir_path)
            return date_hour_dir_path
        except Exception as e:
            print("创建目录 %s 时候，出现异常： %s" %(dir_path,e))
    else:
        return date_hour_dir_path

def create_dir(dir_path):
    try:
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
            return True
        else:
            return True
    except Exception as e:
        print("创建目录 %s 时候，出现异常： %s" %(dir_path,e))
        return False

if __name__ == "__main__":
    create_date_dir("D:\\python_pageobject\\capture_pics")
    create_date_hour_dir("D:\\python_pageobject\\capture_pics")
    create_dir("D:\\python_pageobject\\capture_pic")


