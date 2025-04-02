import pandas as pd
from docx import Document
import os
from pathlib import Path
from datetime import datetime

class ReportGenerator:
    def __init__(self, input_dir="../Inputs", output_dir="../Outputs"):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        
        self.input_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)

    def load_excel_data(self, excel_file):
        file_path = self.input_dir / excel_file
        if not file_path.exists():
            raise FileNotFoundError(f"Excel file not found at {file_path}")
        
        df = pd.read_excel(file_path)
        df.columns = [col.strip('{}') for col in df.columns]
        return df

    def load_template(self, template_file):
        template_path = self.input_dir / template_file
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found at {template_path}")
        
        return Document(template_path)

    def replace_fields(self, document, data_row):
        doc = Document()
        for element in document.element.body:
            doc.element.body.append(element)
            
        for paragraph in doc.paragraphs:
            for key, value in data_row.items():
                str_value = str(value) if pd.notna(value) else ""
                placeholder = f"{{{key}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str_value)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data_row.items():
                        str_value = str(value) if pd.notna(value) else ""
                        placeholder = f"{{{key}}}"
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str_value)
        
        return doc

    def generate_reports(self, excel_file, template_file):
        df = self.load_excel_data(excel_file)
        template = self.load_template(template_file)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        for index, row in df.iterrows():
            report_doc = self.replace_fields(template, row)
            output_filename = f"report_{timestamp}_{index + 1}.docx"
            output_path = self.output_dir / output_filename
            
            report_doc.save(output_path)
            print(f"Generated report: {output_path}")
        
        return f"Generated {len(df)} reports successfully"

def get_file_selection(directory, extension):
    files = [f for f in os.listdir(directory) if f.endswith(extension)]
    if not files:
        return None
    
    print(f"\nAvailable {extension} files:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choice = int(input(f"Select a file (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
            print("Invalid selection. Try again.")
        except ValueError:
            print("Please enter a number.")

def main():
    try:
        generator = ReportGenerator()
        
        print("Select an Excel file for data:")
        excel_file = get_file_selection(generator.input_dir, '.xlsx')
        if not excel_file:
            raise FileNotFoundError("No Excel files found in Inputs directory")
        
        print("\nSelect a Word template file:")
        template_file = get_file_selection(generator.input_dir, '.docx')
        if not template_file:
            raise FileNotFoundError("No Word template files found in Inputs directory")
        
        print(f"\nSelected Excel file: {excel_file}")
        print(f"Selected template: {template_file}")
        confirm = input("Proceed with these files? (y/n): ").lower()
        
        if confirm == 'y':
            result = generator.generate_reports(excel_file, template_file)
            print(result)
        else:
            print("Operation cancelled.")
            
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
