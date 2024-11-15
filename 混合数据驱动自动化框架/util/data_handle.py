def remove_none_from_arr(arr):  # 放到util
    for i in range(len(arr)):
        for j in range(len(arr[i])):
            if arr[i][j] is None:
                arr[i][j] = ""
    return arr

def convert_dict_to_arr(test_data_dict):
    arr = []
    arr.append(list(test_data_dict.keys()))  # 把表头的数据作为第一行
    arr.append(list(test_data_dict.values()))  # 把表头的数据作为第一行
    return arr
