import time

def get_date():
    return str(time.localtime().tm_year)+"-"+str(time.localtime().tm_mon)+"-"+str(time.localtime().tm_mday)

def get_time():
    return str(time.localtime().tm_hour)+":"+str(time.localtime().tm_min)+":"+str(time.localtime().tm_sec)

def get_date_time():
    return get_date()+" "+get_time()

def get_chinese_date():
    return str(time.localtime().tm_year) + "年" + str(time.localtime().tm_mon) + "月" + str(time.localtime().tm_mday)+"日"

def get_chinese_time():
    return str(time.localtime().tm_hour)+"时"+str(time.localtime().tm_min)+"分"+str(time.localtime().tm_sec)+"秒"

def get_chinese_date_time():
    return get_chinese_date()+" "+get_chinese_time()

def get_chinese_hour():
    return str(time.localtime().tm_hour) + "时"

def get_chinese_min():
    return str(time.localtime().tm_min) + "分"

if __name__=="__main__":
    print(get_date())
    print(get_time())
    print(get_date_time())
    print(get_chinese_date())
    print(get_chinese_time())
    print(get_chinese_date_time())
    print(get_chinese_hour())

    print(get_chinese_min())

