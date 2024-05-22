import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_tables()
            tables.extend(table)
    return tables

def save_tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        for i, table in enumerate(tables):
            column_names = [reverse_text(name) for name in table[0]] if i == 0 else table[0]
            if i == 1 and len(table) >= 4:  # Assuming the second table is at index 1 and has at least 4 rows
                data = []
                for row in table[1:]:
                    for val in row:
                        data.append({
                            column_names[0]: val.split('\n')[0],
                            column_names[1]: float(val.split('\n')[1].rstrip('%')) / 100,
                            column_names[2]: float(val.split('\n')[2].rstrip('%')) / 100,
                            column_names[3]: float(val.split('\n')[3].rstrip('%')) / 100
                        })
                df = pd.DataFrame(data)
            else:
                df = pd.DataFrame(table[1:], columns=column_names)
            df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

def reverse_text(text):
    if text:
        return text[::-1]  # Reverse the text
    else:
        return None

if __name__ == "__main__":
    pdf_path = 'PDF_Data.pdf'  # Replace 'your_pdf_file.pdf' with the actual path to your PDF file
    excel_path = 'firstXL.xlsx'  # Specify the output Excel file path

    tables = extract_tables_from_pdf(pdf_path)
    save_tables_to_excel(tables, excel_path)