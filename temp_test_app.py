import sys
sys.path.insert(0, r'c:\Users\Masah\Desktop\目視確認アシスト')
from animal_assistant.workbook import AnimalWorkbook

try:
    wb = AnimalWorkbook.open_batch(r'c:\Users\Masah\Desktop\目視確認アシスト\インプットデータ例\260721')
    print("Success!")
    print(f"Device: {wb.device_number}, Date: {wb.analysis_date}")
    print(f"Total Rows: {len(wb.rows)}")
    row_data = wb.row_data(0)
    print(f"Row 0 Image exists: {row_data['imageExists']}")
    print(f"Row 0 Image Path: {row_data['imagePath']}")
    print(f"Row 0 path error: {row_data['pathError']}")
except Exception as e:
    import traceback
    traceback.print_exc()
