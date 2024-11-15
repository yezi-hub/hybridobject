import os
"""管理各种工程中使用的各种文件路径以及excel中的某些列号"""
#获取当前文件的绝对路径
#print(__file__)

#获得当前文件所在的目录
#print(os.path.dirname(__file__))

#获取当前工程的绝路路径
#print(os.path.dirname(os.path.dirname(__file__)))

# 获得当前文件所在的目录
proj_path = os.path.dirname(os.path.dirname(__file__))

# ini文件的路径
ini_file_path = os.path.join(proj_path,"config","PageElementLocator.ini")

#数据文件的路径
test_data_file_path = os.path.join(proj_path,"test_data","126邮箱测试数据.xlsx")
test_data_dir_path = os.path.join(proj_path,"test_data")

#错误截图的路径
error_capture_pics_dir_path = os.path.join(proj_path,"error_capture_pics")

#报告目录路径
report_dir_path = os.path.join(proj_path,"report")
#报告文件html模板文件路径
template_file_path = os.path.join(proj_path,"util","test.html")

test_case_sn_col_no = 0
test_case_name_col_no = 1
test_case_sheet_name_col_no =3
test_case_test_data_sheet_name_col_no=4
test_case_executed_flag_col_no = 5
test_case_executed_time_col_no = 6
test_case_test_result_col_no = 7
test_case_exception_info_col_no = 8
test_case_capture_pic_path_col_no = 9

test_step_sn_col_no = 0
test_step_keyword_col_no = 2
test_step_locate_method_col_no = 3
test_step_locate_exp_col_no = 4
test_step_element_value_col_no = 5
test_step_executed_time_col_no = 6
test_step_test_result_col_no = 7
test_step_exception_info_col_no = 8
test_step_capture_pic_path_col_no = 9

if __name__=="__main__":
    print(proj_path)
    print(ini_file_path)
    print(test_data_file_path)
    print(error_capture_pics_dir_path )
    print(report_dir_path)
    print(template_file_path)
