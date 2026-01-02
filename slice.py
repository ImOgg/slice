import pandas as pd
import os
import re

# 設定當前資料夾
input_folder = "."  # 代表當前資料夾
file_name = "fish_boat_permission_table.csv"  # 檔案名稱（支援 .csv 或 .xlsx）

# 讀取檔案（自動識別格式）
file_path = os.path.join(input_folder, file_name)
file_extension = os.path.splitext(file_name)[1].lower()

if file_extension == '.csv':
    df = pd.read_csv(file_path, low_memory=False)
    print(f"讀取 CSV 檔案: {file_name}")
elif file_extension in ['.xlsx', '.xls']:
    df = pd.read_excel(file_path, engine='openpyxl')
    print(f"讀取 Excel 檔案: {file_name}")
else:
    raise ValueError(f"不支援的檔案格式: {file_extension}，請使用 .csv 或 .xlsx 檔案")

# 清理 Excel 不支援的非法字元
def clean_illegal_chars(val):
    if isinstance(val, str):
        # 移除 Excel 不支援的控制字元 (0x00-0x1F，除了 tab, newline, carriage return)
        return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', val)
    return val

# 對所有欄位套用清理函數
df = df.map(clean_illegal_chars)
# 取得標題與資料
header = df.columns.tolist()
num_rows = len(df)
split_size = num_rows // 4  # 計算每份的大小
remainder = num_rows % 4  # 計算剩餘的行數

start_idx = 0
for i in range(4):
    # 計算切割範圍
    end_idx = start_idx + split_size + (1 if i < remainder else 0)
    df_part = df.iloc[start_idx:end_idx]  # 取出對應區間的資料

    # 檔案名稱
    output_file = os.path.join(input_folder, f"Excel 檔案名稱_part_{i+1}.xlsx")

    # 寫入 Excel（保留標題）
    df_part.to_excel(output_file, index=False, header=header, engine='openpyxl')

    # 更新下一個區間的起始索引
    start_idx = end_idx

print("Excel 檔案已分割完成！")
