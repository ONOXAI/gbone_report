import os
from parser import parse_xml
from writer import write_excel_with_chart

def main():
    input_folder = "xml_inputs"   # Folder with your XMLs
    output_folder = "reports"     # Folder to save XLSX files

    os.makedirs(output_folder, exist_ok=True)

    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".xml"):
            input_path = os.path.join(input_folder, filename)
            output_name = f"{os.path.splitext(filename)[0]}_report.xlsx"
            output_path = os.path.join(output_folder, output_name)

            try:
                df, metadata = parse_xml(input_path)
                write_excel_with_chart(df, metadata, output_path)
                print(f"✅ Generated: {output_name}")
            except Exception as e:
                print(f"❌ Failed on {filename}: {e}")

if __name__ == "__main__":
    main()
