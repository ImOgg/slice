import pandas as pd
import os

# 設定當前資料夾（因為 knife.py 和 Excel 在同一個資料夾）
input_folder = "."  # 代表當前資料夾
file_name = "Excel 檔案名稱.xlsx"  # Excel 檔案名稱

# 讀取 Excel 檔案
file_path = os.path.join(input_folder, file_name)
df = pd.read_excel(file_path)

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
    df_part.to_excel(output_file, index=False, header=header)

    # 更新下一個區間的起始索引
    start_idx = end_idx

print("Excel 檔案已分割完成！")
