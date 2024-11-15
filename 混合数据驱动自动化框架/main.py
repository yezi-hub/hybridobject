from key_word.key_word import *
from util.excel_util import Excel
from config.var_config import *
from util.time_util import get_chinese_time
from util.capture_pic import capture_pic
from util.generate_report import generate_table_content, gen_html_report, write_html_summary_line
from util.dir_util import create_date_hour_dir
from util.ini_file_parser import IniFileParser
import os, re
import queue, threading, multiprocessing
from util.excel_util import validate_excel_and_sheet
from util.ini_file_parser import get_section_and_option
from util.data_handle import convert_dict_to_arr

def gen_command(step_function_name, locate_method, locate_exp, value):
    if not locate_method and not locate_exp and not value:
        command = step_function_name + "()"
    elif locate_method and locate_exp and not value:
        command = step_function_name + "('%s','%s')" % (locate_method, locate_exp)
    elif not locate_method and not locate_exp and value:
        command = step_function_name + "('%s')" % value
    elif locate_method and locate_exp and value:
        command = step_function_name + "('%s','%s','%s')" % (locate_method, locate_exp, value)
    return command

def execute_test_step(command, test_data_execel_wb,data_dict=None):
    global driver
    test_step_result = "成功"
    exception_info = ""
    pic_path = ""
    try:
        if "open_browser" in command:
            driver = eval(command)
        elif "key_word" in command.lower():
            sheet_name = re.search(r"key_word\('(.*?)'\)",command).group(1)
            if sheet_name in test_data_execel_wb.get_sheet_names():
                driver, test_case_result, test_case_exception_info, pic_path = execute_test_case_by_sheet_name(
                    test_data_execel_wb, sheet_name,data_dict)
        else:
            eval(command)
    except Exception as e:
        exception_info = traceback.format_exc()
        print("执行command: %s 命令的时候，出现异常，异常信息：%s \n %s" % (command, e, exception_info))
        exception_info = "执行command: %s 命令的时候，出现异常，异常信息：%s \n %s" % (
            command, e, exception_info)
        test_step_result = "失败"

        if not isinstance(driver, str):
            pic_path = capture_pic(driver)

    return driver, test_step_result, exception_info, pic_path

def process_value_by_regular_expression(value, test_data_dict):
    try:
        if value and "${" in str(value):
            if re.search(r"\$\{(.*?)\}", value):
                old_var_name = value
                var_name = re.search(r"\$\{(.*?)\}", value).group(1)
                value = test_data_dict[var_name]
                # all_test_steps[test_no][test_step_element_value_col_no] = value
                print("*******", old_var_name, value, test_data_dict)
    except Exception as e:
        print("替换变量 %s 的时候出现异常！异常信息：%s" % (value, e))

    return value

def execute_test_case_by_sheet_name(test_data_execel_wb, test_step_sheet_name,data_dict=None):
    validate_excel_and_sheet(test_data_execel_wb, test_step_sheet_name)
    test_data_execel_wb.set_sheet(test_step_sheet_name)
    all_test_steps = test_data_execel_wb.get_all_rows_values()
    test_step_header = all_test_steps[0]  # 获取到测试步骤sheet的表头
    for test_no in range(1, len(all_test_steps)):
        step_function_name = all_test_steps[test_no][test_step_keyword_col_no]
        locate_method = all_test_steps[test_no][test_step_locate_method_col_no]
        locate_exp = all_test_steps[test_no][test_step_locate_exp_col_no]
        try:
            locate_method, locate_exp = get_section_and_option(ini_file_path, locate_method, locate_exp)
        except Exception as e:
            print(
                "从配置文件读取 section_name：%s option_name:%s 的定位表达式时出现异常" % (locate_method, locate_exp, e))
            raise e

        value = all_test_steps[test_no][test_step_element_value_col_no]
        print("-------------&&&&&&&&&&&&&&&",value,data_dict)
        if data_dict:
            value = process_value_by_regular_expression(value, data_dict)
            all_test_steps[test_no][test_step_element_value_col_no]=value

        executed_time = get_chinese_time()
        all_test_steps[test_no][test_step_executed_time_col_no] = executed_time
        command = gen_command(step_function_name, locate_method, locate_exp, value)
        print(command)
        driver,test_case_result, test_case_exception_info, pic_path = execute_test_step(command,test_data_execel_wb,data_dict)
        all_test_steps[test_no][test_step_exception_info_col_no] = test_case_exception_info
        all_test_steps[test_no][test_step_capture_pic_path_col_no] = pic_path
        all_test_steps[test_no][test_step_test_result_col_no] = test_case_result

    test_data_execel_wb.set_sheet("测试结果")
    test_data_execel_wb.write_lines(all_test_steps, header_color="green")
    global html_report_file_path
    gen_html_report(html_report_file_path, all_test_steps)
    return driver, test_case_result, test_case_exception_info, pic_path

def execute_test_case_by_hybrid(test_data_execel_wb, test_step_sheet_name, test_data_sheet_name):
    driver = ""
    validate_excel_and_sheet(test_data_execel_wb, test_step_sheet_name)
    validate_excel_and_sheet(test_data_execel_wb, test_data_sheet_name)

    # 取出字典的测试数据
    test_data_dict_arr = convert_test_data_format(test_data_execel_wb, test_data_sheet_name)
    if not test_data_dict_arr:  # 如果没有取出任何测试数据，则不运行
        print("测试数据sheet %s 中没有任何测试数据需要运行" % test_data_sheet_name)
        return None

    run_case_num = 0
    for test_data_dict in test_data_dict_arr:  # 如果所有的数据没有一个y的状态，则也不进行运行！
        if test_data_dict["是否执行"] and test_data_dict["是否执行"].lower() == "y":
            run_case_num += 1

    if run_case_num == 0:
        print("测试数据sheet %s 中没有任何测试数据需要运行" % test_data_sheet_name)
        return None

    test_case_result = "成功"
    pic_path = ""
    test_case_exception_info = ""

    for test_data_dict in test_data_dict_arr:
        if "y" not in str(test_data_dict["是否执行"]):
            continue
        test_data_dict["执行时间"] = get_chinese_time()
        pic_path = ""

        driver, test_case_result, test_case_exception_info, pic_path = execute_test_case_by_sheet_name(test_data_execel_wb, test_step_sheet_name, test_data_dict)
        test_data_dict["执行结果"]=test_case_result
        test_data_dict["异常信息"] =test_case_exception_info
        test_data_dict["截图信息"]=pic_path

        test_data_execel_wb.set_sheet("测试结果")
        test_data_execel_wb.write_a_line(list(test_data_dict.keys()), fill="green")
        test_data_execel_wb.write_a_line(list(test_data_dict.values()))
        test_data_arr = convert_dict_to_arr(test_data_dict)
        gen_html_report(html_report_file_path, test_data_arr)
    return driver, test_case_result, test_case_exception_info, pic_path

def execute_test_case_by_file(test_data_file_path):
    hour_dir = create_date_hour_dir(report_dir_path)
    time_file_path = get_chinese_time() + ".html"
    global html_report_file_path
    html_report_file_path = os.path.join(hour_dir, time_file_path)
    test_data_wb = Excel(test_data_file_path)
    test_data_wb.set_sheet("测试用例")
    driver = ""
    test_case_datas = test_data_wb.get_all_rows_values()
    # for test_case_data  in test_case_datas:
    #    print(test_case_data)
    test_case_header = test_case_datas[0]
    # 记录测试用例是否成功的标志
    # for test_case_data in test_case_datas[1:]:
    for i in range(1, len(test_case_datas)):
        test_case_result = "成功"
        # 获得当前用例的执行时间
        test_case_executed_time = get_chinese_time()
        test_case_datas[i][test_case_executed_time_col_no] = test_case_executed_time
        test_case_exception_info = ""
        if test_case_datas[i][test_case_executed_flag_col_no] and "y" in test_case_datas[i][
            test_case_executed_flag_col_no].lower():
            test_data_sheet_name = test_case_datas[i][test_case_test_data_sheet_name_col_no]
            test_case_step_sheet_name = test_case_datas[i][test_case_sheet_name_col_no]
            if not test_data_sheet_name:  # 没有测试数据的sheet名称，则使用关键字的方式来运行
                driver, test_case_result, test_case_exception_info, pic_path = execute_test_case_by_sheet_name(
                    test_data_wb, test_case_step_sheet_name)
            else:  # 有测试数据的sheet，我们使用混合模式去运行
                driver, test_case_result, test_case_exception_info, pic_path = execute_test_case_by_hybrid(test_data_wb,
                                                                                                           test_case_step_sheet_name,
                                                                                                           test_data_sheet_name)
            test_case_datas[i][test_case_test_result_col_no] = test_case_result
            test_case_datas[i][test_case_exception_info_col_no] = test_case_exception_info
            test_case_datas[i][test_case_capture_pic_path_col_no] = pic_path
            test_data_wb.set_sheet("测试结果")
            test_data_wb.write_a_line(test_case_header, fill="green")
            test_data_wb.write_a_line(test_case_datas[i])
            gen_html_report(html_report_file_path, test_case_datas)

    success_case_count = 0
    fail_case_count = 0
    case_count = 0
    for test_case_data in test_case_datas:
        if test_case_data[test_case_test_result_col_no] and "成功" in test_case_data[test_case_test_result_col_no]:
            success_case_count += 1
        elif test_case_data[test_case_test_result_col_no] and "失败" in test_case_data[test_case_test_result_col_no]:
            fail_case_count += 1

    case_count = success_case_count + fail_case_count
    test_data_wb.set_sheet("测试结果")
    test_data_wb.write_a_line(
        [f"用例总数:{case_count}", f"成功用例总数：{success_case_count}", f"失败用例总数：{fail_case_count}"],
        fill="blue")
    data = f"用例总数:{case_count} </br>" + f"成功用例总数：{success_case_count}</br>" + f"失败用例总数：{fail_case_count}</br>"
    write_html_summary_line(html_report_file_path, data)

    test_data_wb.save()


def execute_test_case_by_dir(test_data_dir_path):
    print(test_data_dir_path)
    for i in os.listdir(test_data_dir_path):
        test_data_file_path = os.path.join(test_data_dir_path, i)
        execute_test_case_by_file(test_data_file_path)

def convert_test_data_format(test_data_wb, sheet_name):
    try:
        if test_data_wb and sheet_name in test_data_wb.get_sheet_names():
            test_data_wb.set_sheet(sheet_name)
            test_data = test_data_wb.get_all_rows_values()
            test_data_dict_arr = []  # 用于存储测试数据，每一行数据用一个字典来保存
            for row in range(1, len(test_data)):
                # print("-----",test_data[row])
                data_dict = {}
                for col in range(len(test_data[row])):
                    data_dict[test_data[0][col]] = test_data[row][col]
                test_data_dict_arr.append(data_dict)
        return test_data_dict_arr
    except Exception as e:
        print("转换 %s sheet 的测试数据出现异常！异常信息:%s" % (sheet_name, e))
        return None

def task(queue):
    while not queue.empty():
        test_data_dir_path = queue.get()
        execute_test_case_by_dir(test_data_dir_path)

def concurrent_execute_test_case_dirs(test_case_data_dir):
    if not os.path.exists(test_case_data_dir):
        print("测试用例的目录 %s 不存在，无法并发执行！" % test_case_data_dir)
        return

    dir_path = ""
    test_case_dir_queue = multiprocessing.Queue()
    for i in os.listdir(test_case_data_dir):
        # print(test_case_data_dir,i)
        if os.path.isdir(os.path.join(test_case_data_dir, i)):
            dir_path = os.path.join(test_case_data_dir, i)
            # print(dir_path)
            test_case_dir_queue.put(dir_path)

    # 创建多个线程
    processes = []
    for i in range(3):  # 创建 3 个线程
        process = multiprocessing.Process(target=task, args=(test_case_dir_queue,))
        processes.append(process)
        process.start()

    # 等待所有线程结束
    for process in processes:
        process.join()

if __name__ == "__main__":
    # execute_test_case_by_dir(test_data_dir_path)
    # test_data_wb = Excel(test_data_file_path)
    # print(convert_test_data_format(test_data_wb, "测试数据"))
    concurrent_execute_test_case_dirs(r"d:\hybrid_test_driven_framework\test_data")
