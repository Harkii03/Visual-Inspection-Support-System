import sys
import openpyxl

wb = openpyxl.load_workbook(r'c:\Users\Masah\Desktop\目視確認アシスト\インプットデータ例\260721\result_003003_260721_all.xlsx', data_only=True)
sheet = wb.active

with open('temp_output.txt', 'w', encoding='utf-8') as f:
    for i in range(1, 5):
        f.write(f"Row {i}: {[cell.value for cell in sheet[i]]}\n")
