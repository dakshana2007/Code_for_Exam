from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ===== CONFIG =====
file_path = r"C:\Users\Daksh\OneDrive\Desktop\MYFiles\Attendance_sheet\Clean_Attendance_Sheet.xlsx"
output_path = r"C:\Users\Daksh\OneDrive\Desktop\MYFiles\Attendance_sheet\Final_Clean_Attendance_Sheet.xlsx"
# ==================

# Load workbook
wb = load_workbook(file_path)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"✅ Processing {sheet_name}")

    # Delete columns E, F, G, I
    # Current columns:
    # A B C D E F G H I J
    # 1 2 3 4 5 6 7 8 9 10
    # Delete in reverse order to maintain correct indexing:
    ws.delete_cols(9)  # I
    ws.delete_cols(7)  # G
    ws.delete_cols(6)  # F
    ws.delete_cols(5)  # E

    # Remove text wrap from column C ("Dakshana Roll No -- Name")
    col_letter = 'C'
    for cell in ws[col_letter]:
        current_alignment = cell.alignment
        ws[cell.coordinate].alignment = Alignment(
            horizontal=current_alignment.horizontal if current_alignment else None,
            vertical=current_alignment.vertical if current_alignment else None,
            wrap_text=False
        )

# Save modified workbook
wb.save(output_path)

print(f"\n✅ Final cleaned Excel saved at:\n{output_path}")
