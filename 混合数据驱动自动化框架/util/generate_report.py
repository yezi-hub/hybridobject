#data需要是一个二维列表
from config.var_config import *
from util.file_util import read_file,append_file
from util.data_handle import remove_none_from_arr

def generate_table_content(*data):
    template = """  
    <table>  
    <thead>  
        <tr>  
            {head_row}  
        </tr>  
    </thead>  
    <tbody>  
        {table_rows}  
    </tbody>  
    </table>   
    <br><br>  
    """
    content = ""

    for table_data in data:
        if not table_data:
            continue  # 跳过空的表格数据

        head_row = table_data[0]
        print("表头：", head_row)

        # 生成表头
        head_cells = "".join(f"<th>{cell}</th>" for cell in head_row)

        # 生成其他行
        other_rows = ""
        for row_data in table_data[1:]:
            # 生成每行的单元格
            row_cells = "".join(f"<td>{cell}</td>" for cell in row_data)
            other_rows += f"<tr>{row_cells}</tr>\n"

            # 格式化整个表格
        table = template.format(head_row=head_cells, table_rows=other_rows)
        content += table

    return content

def write_html_summary_line(html_report_file_path, data):
    template = """  
        <table>  
        <thead>  
            <tr >  
                <td class="header-text">%s</d>
            </tr>  
        </thead>  

        </table>   
        <br><br>  
        """
    html = template % data
    with open(html_report_file_path, "a", encoding="utf-8") as fp:
        fp.write(html + "\n")

def gen_html_report(html_report_file_path, test_report_data):
    remove_none_from_arr(test_report_data)
    table_content = generate_table_content(test_report_data)
    template_html = read_file(template_file_path)
    try:
        if not os.path.exists(html_report_file_path):
            report_file_html = ""
        else:
            report_file_html = read_file(template_file_path)
    except Exception as e:
        report_file_html = ""
    #加锁，写html报告，避免同时写入有冲突
    #lock = multiprocessing.Lock()
    #with lock:
    try:
        if "<style>" in report_file_html or "自动化测试报告" in report_file_html:
            append_file(html_report_file_path, "\n" + table_content)
        else:
            append_file(html_report_file_path,template_html + table_content)
    except Exception as e:
        print("生成html报告 %s 文件失败，异常信息：%s" % (html_report_file_path, e))

if __name__ == "__main__":
    from util.excel_util import  Excel
    from config.var_config import test_data_file_path
    test_report_wb = Excel(test_data_file_path)
    test_report_wb.set_sheet("测试结果")
    test_report_data = test_report_wb.get_all_rows_values()
    for i in range(len(test_report_data)):
        for j in range(len(test_report_data[i])):
            if test_report_data[i][j] is None:
                test_report_data[i][j]=""
    print("test_data:",test_report_data)
    content = generate_table_content(test_report_data,test_report_data)


    with open(template_file_path,encoding="utf-8") as fp:
        html = fp.read()

    with open("e:\\report.html", "a+", encoding="utf-8") as fp:
        if "<style>" in fp.read():
            fp.write("\n"+content)
        else:
            fp.write(html + content)