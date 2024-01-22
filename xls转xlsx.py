#1
import os
import glob
import pandas as pd

def xls_to_xlsx(input_path, output_path):
    # 读取xls文件
    xls_data = pd.read_excel(input_path, None)

    # 保存为xlsx文件
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, df in xls_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def convert_xls_to_xlsx_in_directory(directory):
    # 获取目录下所有的xls和xlsx文件
    xls_files = glob.glob(os.path.join(directory, '*.xls'))
    xlsx_files = glob.glob(os.path.join(directory, '*.xlsx'))

    all_files = xls_files + xlsx_files

    for xls_file in xls_files:
        # 获取xls文件的路径和文件名
        xls_path = os.path.abspath(xls_file)
        xlsx_path = os.path.splitext(xls_path)[0] + '.xlsx'  # 将文件后缀改为xlsx

        # 转换xls到xlsx
        xls_to_xlsx(xls_path, xlsx_path)

        # 删除原始的xls文件
        os.remove(xls_path)

if __name__ == "__main__":
    # 指定目录
    target_directory = r'K:\医保局2024\职工大额格式\excel数据职工大额test'

    # 调用函数进行转换
    convert_xls_to_xlsx_in_directory(target_directory)
