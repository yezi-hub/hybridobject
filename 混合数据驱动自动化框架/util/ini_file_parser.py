import configparser
import os

class IniFileParser:

    def __init__(self,ini_file_path):
        if not os.path.exists(ini_file_path):
            print("ini文件 %s 的路径不存在！" %ini_file_path)
            self.ini_file_path = None
            return
        self.ini_file_path = ini_file_path
        # 创建一个配置解析器
        self.config = configparser.ConfigParser()
        # 读取 UTF-8 编码的 INI 文件
        with open(self.ini_file_path, 'r', encoding='utf-8') as config_file:
            self.config.read_file(config_file)



    #获取当前类中存储的ini文件的方法
    def get_ini_file_path(self):
        return self.ini_file_path

    def set_ini_file_path(self,ini_file_path):
        if not os.path.exists(ini_file_path):
            print("ini文件 %s 的路径不存在！" %ini_file_path)
            self.ini_file_path = None
            return
        self.ini_file_path = ini_file_path
        self.config.read(ini_file_path)#读取配置文件的配置信息

    def get_option_value(self,section_name,option_name):
        if self.config.has_option(section_name, option_name):
            return self.config[section_name][option_name]
        return None

def get_section_and_option(ini_file_path,locate_method, locate_exp):
    # 满足if条件的话，按照读取section_name和option_name去处理
    if locate_method and locate_exp and (
            locate_method not in ["id", "name", "tag_name", "partial_link_text", "link_text", "xpath"]):
        section_name = locate_method.strip()
        option_name = locate_exp.strip()
        ini_parser = IniFileParser(ini_file_path)
        locate_info = ini_parser.get_option_value(section_name, option_name)
        if not locate_info:
            print("section_name：%s option_name:%s 没有在配置文件中读到对应值：" % (section_name, option_name))
            raise Exception("section_name：%s option_name:%s 没有在配置文件中读到对应值：" % (section_name, option_name))
        else:
            locate_method = locate_info.split("||")[0].strip()
            locate_exp = locate_info.split("||")[1].strip()
    return locate_method, locate_exp

if __name__ == "__main__":
    ini_parser = IniFileParser("e:\\pytest.ini")
    print(ini_parser.get_option_value("pytest","python_files"))