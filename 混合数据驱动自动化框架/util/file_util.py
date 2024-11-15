import os

def read_file(file_path, encoding="utf-8"):
    if not os.path.exists(file_path):
        print("读取的文件路径 %s 不存在！" % (file_path))
        raise FileExistsError("读取的文件路径 %s 不存在！" % (file_path))
    try:
        with open(file_path, encoding=encoding) as fp:
            content = fp.read().strip()
            return content
    except Exception as e:
        print("读取的文件 %s 的时候出现异常，异常信息：%s" % (file_path, e))
        raise Exception("读取的文件 %s 的时候出现异常，异常信息：%s" % (file_path, e))

def append_file(file_path, content, encoding="utf-8"):
    try:
        with open(file_path, "a", encoding=encoding) as fp:
            content = fp.write(content)
    except Exception as e:
        print("文件 %s 追加内容的时候出现异常，异常信息：%s" % (file_path, e))
        raise Exception("文件 %s 追加内容的时候出现异常，异常信息：%s" % (file_path, e))