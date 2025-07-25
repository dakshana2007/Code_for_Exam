import os
import pandas as pd

# === Step 1: Set file paths ===
excel_file = r'C:\Users\Daksh\OneDrive\Desktop\MYFiles\Re_Admit\Rename_admit.xlsx'
pdf_folder = r'C:\Users\Daksh\OneDrive\Desktop\MYFiles\Re_Admit\Admit_card'
column_name = 'NewName'  # Name of the column in Excel

# === Step 2: Load Excel ===
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    print(f"Excel file not found at: {excel_file}")
    exit()

# === Step 3: Clean names ===
new_names = df[column_name].astype(str).str.strip().tolist()

# === Step 4: Load PDF files sorted by creation time ===
try:
    pdf_files = sorted(
        [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')],
        key=lambda x: os.path.getctime(os.path.join(pdf_folder, x))
    )
except FileNotFoundError:
    print(f"PDF folder not found at: {pdf_folder}")
    exit()

# === Step 5: Check file count ===
if len(pdf_files) != len(new_names):
    print(f"Mismatch: {len(pdf_files)} PDFs ≠ {len(new_names)} names in Excel.")
    exit()

# === Step 6: Rename files safely ===
for old_file, new_name in zip(pdf_files, new_names):
    old_path = os.path.join(pdf_folder, old_file)
    new_path = os.path.join(pdf_folder, new_name + '.pdf')

    # Skip if already renamed correctly
    if old_file == new_name + '.pdf':
        print(f"⏭️ Already renamed: {old_file}")
        continue

    # Skip renaming if target already exists to avoid overwriting
    if os.path.exists(new_path):
        print(f"⚠️ Skipping (duplicate exists): {new_name}.pdf")
        continue

    try:
        os.rename(old_path, new_path)
        print(f"✅ Renamed: {old_file} → {new_name}.pdf")
    except Exception as e:
        print(f"❌ Error renaming {old_file}: {e}")

print("\n🎉 Renaming completed without deleting any file.")
