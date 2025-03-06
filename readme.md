# Excel 檔案分割工具

這個工具可以將一個 Excel 檔案平均分割成 4 個部分。

## 使用需求

- Python 3.x
- pandas 套件
- openpyxl 套件 (用於處理 .xlsx 檔案)

## 安裝相依套件

```bash
pip install pandas openpyxl
```

## 使用方法

1. 將要分割的 Excel 檔案放在與 slice.py 相同的資料夾中
2. 打開 slice.py，修改以下變數：
   ```python
   file_name = "Excel 檔案名稱.xlsx"  # 改成您的 Excel 檔案名稱
   ```
3. 執行 slice.py：
   ```bash
   python slice.py
   ```

## 輸出結果

程式會在同一個資料夾中生成 4 個新的 Excel 檔案：
- Excel 檔案名稱_part_1.xlsx
- Excel 檔案名稱_part_2.xlsx
- Excel 檔案名稱_part_3.xlsx
- Excel 檔案名稱_part_4.xlsx

每個檔案都會包含原始檔案的標題列，並且資料會被平均分配。