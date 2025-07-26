import sys
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import load_workbook

def extract_kml_from_kmz(file_path):
    with zipfile.ZipFile(file_path, 'r') as z:
        for filename in z.namelist():
            if filename.endswith('.kml'):
                return z.read(filename).decode('utf-8')
    return None

def extract_placemarks(kml_content):
    root = ET.fromstring(kml_content)
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    placemarks = []

    for pm in root.findall('.//kml:Placemark', ns):
        name_elem = pm.find('kml:name', ns)
        coord_elem = pm.find('.//kml:coordinates', ns)

        if name_elem is not None and coord_elem is not None:
            name = name_elem.text.strip()
            coords = coord_elem.text.strip().split(',')

            if len(coords) >= 2:
                lon = round(float(coords[0]), 6)
                lat = round(float(coords[1]), 6)

                if name.isalnum() and not name.isnumeric() and len(name) > 5:
                    splitter_name = name + '01'
                    rsa = name[:3].upper()

                    placemarks.append({
                        'FatBox': name,
                        'splitter name': splitter_name,
                        'rsa': rsa,
                        'latitude': lat,
                        'longitude': lon
                    })

    return placemarks

def auto_fit_excel_columns(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for col in worksheet.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max_length + 2

def process_file(file_path, output_excel):
    if file_path.endswith('.kmz'):
        kml_content = extract_kml_from_kmz(file_path)
    elif file_path.endswith('.kml'):
        with open(file_path, 'r', encoding='utf-8') as f:
            kml_content = f.read()
    else:
        raise ValueError("File must be .kml or .kmz")

    placemarks = extract_placemarks(kml_content)
    df = pd.DataFrame(placemarks)
    df.sort_values(by='FatBox', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.index += 1
    df.reset_index(inplace=True)
    df.rename(columns={'index': 's.no'}, inplace=True)

    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Placemarks')
        auto_fit_excel_columns(writer, 'Placemarks')

    print(f"Extracted and saved {len(df)} placemarks to '{output_excel}'")

# Entry point when called from Node.js
if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python convert.py <input_file_path> <output_excel_path>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    process_file(input_path, output_path)
