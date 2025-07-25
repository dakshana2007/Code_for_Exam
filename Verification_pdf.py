import os
import win32com.client

# Path to your Excel file
excel_file = r"C:\Users\Daksh\OneDrive\Desktop\MYFiles\Verification_list\Generated_Verification_By_Center.xlsx"

# Output folder for the PDFs
output_folder = os.path.dirname(excel_file)

# Launch Excel in the background
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

# Open the workbook
wb = excel.Workbooks.Open(excel_file, ReadOnly=True)

try:
    for sheet in wb.Sheets:
        # Set print area to used range
        used_range = sheet.UsedRange
        sheet.PageSetup.PrintArea = used_range.Address

        # Page setup
        ps = sheet.PageSetup
        ps.Orientation = 1  # Portrait
        ps.PaperSize = 9    # A4
        ps.Zoom = False     # Disable default zoom
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False

        # Reduce margins (in inches)
        ps.LeftMargin = excel.InchesToPoints(0.3)
        ps.RightMargin = excel.InchesToPoints(0.3)
        ps.TopMargin = excel.InchesToPoints(0.3)
        ps.BottomMargin = excel.InchesToPoints(0.3)
        ps.HeaderMargin = excel.InchesToPoints(0)
        ps.FooterMargin = excel.InchesToPoints(0)

        # Center horizontally and vertically
        ps.CenterHorizontally = True
        ps.CenterVertically = True

        # Repeat header row (row 1)
        ps.PrintTitleRows = "$1:$1"

        # Export to PDF
        pdf_path = os.path.join(output_folder, f"{sheet.Name}.pdf")
        sheet.ExportAsFixedFormat(0, pdf_path)

        print(f"✅ Exported: {pdf_path}")

finally:
    wb.Close(False)
    excel.Quit()
