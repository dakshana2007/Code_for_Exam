import os
import re
import pdfplumber

folder_path = r"C:\Users\Daksh\OneDrive\Desktop\Rehearsal DST\Admit card\Admit"

for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):
        file_path = os.path.join(folder_path, filename)
        roll_number = None  # Reset per file

        try:
            with pdfplumber.open(file_path) as pdf:
                first_page = pdf.pages[0]
                words = first_page.extract_words()

                for i, word in enumerate(words):
                    if "Dakshana" in word['text']:
                        for j in range(i + 1, min(i + 4, len(words))):
                            if re.match(r'\d{11}', words[j]['text']):
                                roll_number = words[j]['text']
                                break
                    if roll_number:
                        break

            # ✅ Now the PDF is closed — it's safe to rename
            if roll_number:
                new_file_path = os.path.join(folder_path, roll_number + '.pdf')

                if not os.path.exists(new_file_path):
                    os.rename(file_path, new_file_path)
                    print(f'✅ Renamed: {filename} → {roll_number}.pdf')
                else:
                    print(f'⚠️ Skipped: {roll_number}.pdf already exists.')
            else:
                print(f'❌ Roll number not found in {filename}')

        except Exception as e:
            print(f'⚠️ Error reading {filename}: {e}')
