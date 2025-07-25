import pandas as pd
import math
import xlsxwriter
import os

# --- CONFIG ---
input_file = r'C:\Users\Daksh\OneDrive\Desktop\MYFiles\Data_For_all\Data_Student.xlsx'
output_file = r'C:\Users\Daksh\OneDrive\Desktop\MYFiles\Seat_allotment\Seat_Allotment_Result_Final.xlsx'
rows, cols = 5, 4  # Seating config
seats_per_room = rows * cols

# --- CHECK IF INPUT EXISTS ---
if not os.path.exists(input_file):
    print(f"❌ File not found at: {input_file}")
    exit()

# --- LOAD & CLEAN DATA ---
df = pd.read_excel(input_file)
df = df[['Dakshana Roll No -- Name', 'Test Center']]
df.dropna(subset=['Dakshana Roll No -- Name'], inplace=True)
df.drop_duplicates(subset=['Dakshana Roll No -- Name'], inplace=True)
df.reset_index(drop=True, inplace=True)

# --- CREATE WORKBOOK ---
workbook = xlsxwriter.Workbook(output_file)

# --- FORMATS ---
title_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'font_size': 16, 'border': 0
})

row_number_header_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'border': 2,  'font_size': 12,
})

grid_format = workbook.add_format({
    'align': 'left', 'valign': 'vcenter', 'border': 2, 'indent': 1, 'text_wrap': True
})

blank_format = workbook.add_format({
    'italic': True, 'align': 'center', 'valign': 'vcenter',
    'border': 2, 'font_color': 'gray',  'font_size': 16,
})

side_header_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'border': 2, 'rotation': 90,  'font_size': 16,
})

row_header_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'border': 2,  'font_size': 16,
})

column_number_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'border': 2,  'font_size': 12,
})
header_label_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'font_size': 16, 'border': 2
})

header_number_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'font_size': 12, 'border': 2
})

side_header_label_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'font_size': 16, 'border': 2, 'rotation': 90
})

try:
    for center_name, center_df in df.groupby('Test Center'):
        worksheet = workbook.add_worksheet(name=str(center_name))
        students = list(center_df['Dakshana Roll No -- Name'])
        num_rooms = math.ceil(len(students) / seats_per_room)
        row_pointer = 0

        # Set column widths
        worksheet.set_column(0, 0, 5)   # Column A: Column Number (vertical)
        worksheet.set_column(1, 1, 3)   # Column B: Row numbers
        worksheet.set_column(2, cols + 1, 30)  # Columns C-F: Student names

        for room_number in range(num_rooms):
            if room_number == 0:
                row_pointer += 1
            start = room_number * seats_per_room
            end = start + seats_per_room
            room_students = students[start:end]

            while len(room_students) < seats_per_room:
                room_students.append("BLANK")

            # --- Room Headers ---
            worksheet.merge_range(row_pointer, 3, row_pointer, 4, 'Seating Plan for JDST 2026 (For Class 12)', title_format)
            row_pointer += 1
            worksheet.merge_range(row_pointer, 3, row_pointer, 4, f'Center Number : {center_name}', title_format)
            row_pointer += 1
            worksheet.merge_range(row_pointer, 3, row_pointer, 4, f'Room Number : {room_number + 1}', title_format)
            row_pointer += 1

            # --- "Row Number" Header Row ---
            no_border_format = workbook.add_format({'border': 0}) 
            worksheet.write(row_pointer, 0, '', no_border_format)
            worksheet.write(row_pointer, 1, '', no_border_format)
            worksheet.merge_range(row_pointer, 2, row_pointer, cols + 1, 'Row Number', header_label_format)
            worksheet.set_row(row_pointer, 20) 
            row_pointer += 1

            worksheet.write(row_pointer, 0, '', no_border_format)
            worksheet.write(row_pointer, 1, '', no_border_format)
            for c in range(cols):
                worksheet.write(row_pointer, c + 2, str(c + 1), header_number_format)
                worksheet.set_row(row_pointer, 20)
            row_pointer += 1

            worksheet.merge_range(row_pointer, 0, row_pointer + rows - 1, 0, 'Column Number', side_header_label_format)

            for r in range(rows):
                worksheet.set_row(row_pointer + r, 30)
                worksheet.write(row_pointer + r, 1, str(r + 1), header_number_format)

            # --- Seat Grid (C to F) ---
            for r in range(rows):
                for c in range(cols):
                    idx = r * cols + c
                    student = room_students[idx]
                    fmt = blank_format if student == "BLANK" else grid_format
                    worksheet.write(row_pointer + r, c + 2, student, fmt)

            row_pointer += rows + 1

finally:
    workbook.close()
    print(f"✅ Seat allotment saved to: {output_file}")

# --- EXPORT EACH SHEET TO PDF (One PDF per center) ---
# --- EXPORT EACH SHEET TO PDF (One PDF per center) ---
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

wb = excel.Workbooks.Open(output_file)

for sheet in wb.Sheets:
    sheet.PageSetup.Orientation = 2  # Landscape
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = False
    sheet.PageSetup.CenterHorizontally = True
    sheet.PageSetup.CenterVertically = True
    sheet.PageSetup.TopMargin = excel.InchesToPoints(0.3)  # ← ADD THIS LINE

    center_number = sheet.Name
    pdf_path = os.path.join(os.path.dirname(output_file), f"{center_number}.pdf")
    sheet.ExportAsFixedFormat(0, pdf_path)

wb.Close(SaveChanges=False)
excel.Quit()

