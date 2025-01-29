from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from io import BytesIO
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import os
import sys
from xml.dom import minidom

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# ==================================================================
# Shared Conversion Functions
# ==================================================================

COLUMN_MAPPINGS = {
    "Name": "Name",
    "R": "Red",
    "B": "Blue",
    "G": "Green",
    "diffuse": "Matte",
    "reflective": "Specularity",
    "transparent": "Transparency",
    "refractive-index-minus-one": "Refractive Index - 1",
    "luminous": "Luminous",
    "Texture Asset": "Texture File Path",
    "texture-mapping": "Texture Mapping 1= Normal; 2=Cylinder; 3=Sphere",
    "texture-multiply-color": "Blend Texture (1=yes; 0=no)",
    "Size U": "Size U",
    "Size V": "Size V",
    "Bumpmap Asset": "Bump Texture File Path",
    "bumpmap-angle": "Bump Height (0-90)",
    "construction-type": "Category",
    "article-no": "Article Code",
    "description": "Description",
    "Price": "Price",
    "core-material": "Core Material",
    "Weight,spec.(g/cm³)": "Weight",
    "material-anisotropic": "Material has Grain",
    "user-attri1": "Attribute 1",
    "user-attri2": "Attribute 2",
    "user-attri3": "Attribute 3",
    "user-attri4": "Attribute 4",
    "user-attri5": "Attribute 5",
    "user-attri6": "Attribute 6",
    "side-grain": "Side Grain",
    "end-grain": "End Grain"
}

COLUMN_ORDER = [
    "Name", "Red", "Green", "Blue", "Matte", "Specularity", "Transparency",
    "Refractive Index - 1", "Luminous", "Texture File Path",
    "Texture Mapping 1= Normal; 2=Cylinder; 3=Sphere", "Blend Texture (1=yes; 0=no)",
    "Size U", "Size V", "Bump Texture File Path", "Bump Height (0-90)", "Category",
    "Article Code", "Description", "Price", "Core Material", "Weight","Material has Grain",
    "Attribute 1", "Attribute 2", "Attribute 3", "Attribute 4", "Attribute 5", "Attribute 6", "Side Grain", "End Grain"
]

DECIMAL_COLUMNS = ["Red", "Green", "Blue", "Matte", "Specularity", "Transparency", "Refractive Index - 1", "Luminous"]

# ==================================================================
# Excel to XML Conversion
# ==================================================================


def convert_xml_to_excel(xml_content):
    root = ET.fromstring(xml_content)
    namespaces = {
        "": root.tag.split("}")[0].strip("{"),
        "a": "http://xmlns.pytha.com/attributes/1.0"
    }

    materials_data = []
    assets_data = {}
    # Extract asset data
    assets = root.find(".//assets", namespaces)
    if assets is not None:
        for asset in assets.findall("asset", namespaces):
            asset_data = extract_asset_data(asset, namespaces)
            asset_id = asset_data.get("Asset ID")
            asset_url = asset_data.get("Source URL")
            if asset_id and asset_url:
                assets_data[asset_id] = asset_url
    # Extract material data
    materials = root.find(".//materials", namespaces)
    if materials is not None:
        for material in materials.findall("material", namespaces):
            material_data = extract_material_data(material, namespaces)
            # Handle texture asset
            texture_urls = []
            texture_assets = material.find(".//a:texture-asset", namespaces)
            if texture_assets is not None:
                for idx, f in enumerate(texture_assets.findall("a:f", namespaces)):
                    if idx == 1:  # Only second <a:f>
                        asset_id = f.text.strip() if f.text else None
                        if asset_id and asset_id in assets_data:
                            texture_urls.append(assets_data[asset_id])
            material_data["Texture Asset"] = ", ".join(texture_urls) if texture_urls else None
            # Handle bumpmap asset
            bumpmap_urls = []
            bm_assets = material.find(".//a:bumpmap-asset", namespaces)
            if bm_assets is not None:
                for idx, f in enumerate(bm_assets.findall("a:f", namespaces)):
                    if idx == 1:  # Only second <a:f>
                        asset_id = f.text.strip() if f.text else None
                        if asset_id and asset_id in assets_data:
                            bumpmap_urls.append(assets_data[asset_id])
            material_data["Bumpmap Asset"] = ", ".join(bumpmap_urls) if bumpmap_urls else None
            materials_data.append(material_data)
    # Filter and organize columns based on the specified order
    filtered_data = []
    for material in materials_data:
        filtered_row = {}
        for column in COLUMN_ORDER:
            old_key = [key for key, value in COLUMN_MAPPINGS.items() if value == column]
            is_decimal = column in DECIMAL_COLUMNS
            filtered_row[column] = convert_to_numeric(material.get(old_key[0]) if old_key else None, is_decimal)
        # Keep missing values as None
        for column in COLUMN_ORDER:
            if filtered_row[column] is None:
                filtered_row[column] = None
        filtered_data.append(filtered_row)
    # Write to Excel
    output_file = os.path.splitext(os.path.basename(file_path))[0] + "_converted.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Materials"

    # Write headers
    ws.append(COLUMN_ORDER)
    # Write rows
    for material in filtered_data:
        row = [material.get(column) for column in COLUMN_ORDER]
        ws.append(row)

    # Create Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df = pd.DataFrame(filtered_data, columns=COLUMN_ORDER)
        df.to_excel(writer, index=False)
    output.seek(0)
    
    return output

def add_attribute(attributes, tag, id_value, name, value, value_type):
    attrib_params = {
        "id": str(id_value),
        "t": value_type
    }
    if name:
        attrib_params["name"] = name
    ordered_params = ["id", "name", "t"]
    ordered_params = [key for key in ordered_params if key in attrib_params]
    attr = ET.SubElement(attributes, tag, {key: attrib_params[key] for key in ordered_params})
    if value is not None:
        attr.text = str(value)

def create_xml_from_excel(excel_data):
    current_datetime = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
    pymat = ET.Element("pymat", {
        "xmlns": "http://xmlns.pytha.com/materials/1.0",
        "xmlns:a": "http://xmlns.pytha.com/attributes/1.0"
    })

    # Create header element
    header = ET.SubElement(pymat, "header")

    # Add unit-system to header
    unit_system = ET.SubElement(header, "unit-system")
    length = ET.SubElement(unit_system, "length")
    length.text = "mm"

    # Add RDF structure to header
    rdf = ET.SubElement(header, "rdf:RDF", {
        "xmlns:rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
        "xmlns:xmp": "http://ns.adobe.com/xap/1.0/"
    })

    description = ET.SubElement(rdf, "rdf:Description", {"rdf:about": ""})
    creator_tool = ET.SubElement(description, "xmp:CreatorTool")
    creator_tool.text = "PYTHA V25"

    # Add the current date and time in the <xmp:CreateDate> tag
    create_date = ET.SubElement(description, "xmp:CreateDate")
    create_date.text = current_datetime

    # Add extension-attributes, materials, and assets elements with a space inside the tags
    extension_attributes = ET.SubElement(pymat, "extension-attributes")
    extension_attributes.text = "\n  "  # Add a space

    materials = ET.SubElement(pymat, "materials")
    materials.text = ""  # Add a space

    # Loop through each row in the Excel data and create materials
    texture_file_paths = excel_data['Texture File Path'].dropna().unique()
    bump_texture_file_paths = excel_data['Bump Texture File Path'].dropna().unique()

    # Merge both sets of file paths and remove duplicates
    all_file_paths = set(texture_file_paths) | set(bump_texture_file_paths)

    # Create assets block and add all unique file paths
    assets = ET.SubElement(pymat, "assets")
    for asset_id, file_path in enumerate(all_file_paths, start=1):
        asset = ET.SubElement(assets, "asset", {"id": str(asset_id)})
        source = ET.SubElement(asset, "source", {"url": f"{file_path}"})
        attributes = ET.SubElement(asset, "a:attributes")
        attributes.text = " "

    # Now loop through each material
    for index, row in excel_data.iterrows():
        material = ET.SubElement(materials, "material", {
            "name": str(row['Name']),
            "id": str(index + 1)
        })

        # Create attributes block specific to this row
        attributes = ET.SubElement(material, "a:attributes")
        attributes.text = ""  # Ensuring the tag is closed properly

        # Add attributes as per the required structure
        add_attribute(attributes, "a:article-no", 2, "Article no", row['Article Code'], "s")
        add_attribute(attributes, "a:description", 4, "Description", row['Description'], "s")
        add_attribute(attributes, "a:core-material", 14, "Core material", row['Core Material'], "s")
        add_attribute(attributes, "a:construction-type", 45, "Construction type", row.get('Category', ''), "s")
        add_attribute(attributes, "a:user-attri1", 51, "User Attri1", row.get('Attribute 1', ''), "s")
        add_attribute(attributes, "a:user-attri2", 52, "User Attri2", row.get('Attribute 2', ''), "s")
        add_attribute(attributes, "a:user-attri3", 53, "User Attri3", row.get('Attribute 3', ''), "s")
        add_attribute(attributes, "a:user-attri4", 54, "User Attri4", row.get('Attribute 4', ''), "s")
        add_attribute(attributes, "a:user-attri5", 55, "User Attri5", row.get('Attribute 5', ''), "s")
        add_attribute(attributes, "a:user-attri6", 56, "User Attri6", row.get('Attribute 6', ''), "s")
        add_attribute(attributes, "a:attribute", 61, "Weight,spec.(g/cm³)", row['Weight'], "d")
        add_attribute(attributes, "a:material-anisotropic", 126, "Material has grain", row.get('Material has Grain', ''), "i")
        add_attribute(attributes, "a:attribute", 1011, "Price", row['Price'], "s")

        # Add texture-srgb element
        if pd.notna(row['Red']) and pd.notna(row['Green']) and pd.notna(row['Blue']):
            color_srgb = ET.SubElement(attributes, "a:color-srgb", {"id": "1000001", "t": "c"})
            red = ET.SubElement(color_srgb, "a:f", {"r": "0", "t": "d"})
            red.text = str(row['Red'])
            green = ET.SubElement(color_srgb, "a:f", {"r": "0", "t": "d"})
            green.text = str(row['Blue'])
            blue = ET.SubElement(color_srgb, "a:f", {"r": "0", "t": "d"})
            blue.text = str(row['Green'])

        # Additional attributes...
        add_attribute(attributes, "a:diffuse", 1000002, "", row['Matte'], "d")
        add_attribute(attributes, "a:reflective", 1000003, "", row['Specularity'], "d")
        add_attribute(attributes, "a:transparent", 1000004, "", row['Transparency'], "d")
        add_attribute(attributes, "a:luminous", 1000005, "", row['Luminous'], "d")
        add_attribute(attributes, "a:refractive-index-minus-one", 1000006, "", row['Refractive Index - 1'], "d")
        add_attribute(attributes, "a:texture-mapping", 1000008, "", row['Blend Texture (1=yes; 0=no)'], "i")

        # Texture size block
        if pd.notna(row['Size U']) and pd.notna(row['Size V']):
            texture_size = ET.SubElement(attributes, "a:texture-size", {"id": "1000009", "t": "c"})

            size_u = ET.SubElement(texture_size, "a:f", {"r": "0", "t": "d", "i": "4"})
            size_u.text = str(row['Size U'])

            size_v = ET.SubElement(texture_size, "a:f", {"r": "0", "t": "d", "i": "4"})
            size_v.text = str(row['Size V'])

        # Add texture and bumpmap assets with the required sub-elements
        if pd.notna(row['Texture File Path']):
            texture_asset = ET.SubElement(attributes, "a:texture-asset", {"id": "1000010", "t": "c"})
            texture_file_path = row['Texture File Path']
            texture_id = list(all_file_paths).index(texture_file_path) + 1
            
            # Adding the first, second, and third <a:f> sub-elements to texture-asset
            ET.SubElement(texture_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = "33"
            ET.SubElement(texture_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = str(texture_id)  # Set the correct texture ID
            ET.SubElement(texture_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = "0"

        add_attribute(attributes, "a:texture-multiply-color", 1000011, "", row.get('Blend Texture (1=yes; 0=no)', '1'), "i")
        
        if pd.notna(row['Bump Texture File Path']):
            bumpmap_asset = ET.SubElement(attributes, "a:bumpmap-asset", {"id": "1000012", "t": "c"})
            bump_texture_file_path = row['Bump Texture File Path']
            bump_id = list(all_file_paths).index(bump_texture_file_path) + 1
            
            # Adding the first, second, and third <a:f> sub-elements to bumpmap-asset
            ET.SubElement(bumpmap_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = "33"
            ET.SubElement(bumpmap_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = str(bump_id)  # Set the correct bumpmap ID
            ET.SubElement(bumpmap_asset, "a:f", {"r": "0", "t": "i", "i": "1"}).text = "0"

        add_attribute(attributes, "a:bumpmap-angle", 1000013, "", row['Bump Height (0-90)'], "d")

    # Convert to pretty XML
    xml_str = ET.tostring(pymat, encoding="unicode")
    parsed_str = minidom.parseString(xml_str)
    pretty_xml = parsed_str.toprettyxml(indent="  ")
    pretty_xml = "\n".join(pretty_xml.splitlines()[1:])
    pretty_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' + pretty_xml.lstrip()
    
    return pretty_xml.encode('utf-8')

# ==================================================================
# Flask Routes
# ==================================================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def handle_conversion():
    if 'file' not in request.files:
        return "No file uploaded", 400
        
    file = request.files['file']
    conversion_type = request.form.get('conversion_type')
    
    if file.filename == '':
        return "No selected file", 400

    try:
        # Validate file type
        filename = secure_filename(file.filename)
        ext = filename.rsplit('.', 1)[1].lower()
        
        if conversion_type == 'excel2xml' and ext != 'xlsx':
            return "Invalid file type for Excel to XML conversion", 400
        if conversion_type == 'xml2excel' and ext != 'pymat':
            return "Invalid file type for XML to Excel conversion", 400

        # In-memory processing
        file_data = file.read()
        
        if conversion_type == 'excel2xml':
            excel_data = pd.read_excel(BytesIO(file_data))
            xml_content = create_xml_from_excel(excel_data)
            return send_file(
                BytesIO(xml_content),
                mimetype='application/xml',
                as_attachment=True,
                download_name='converted_materials.pymat'
            )
            
        elif conversion_type == 'xml2excel':
            excel_file = convert_xml_to_excel(file_data)
            return send_file(
                excel_file,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='converted_materials.xlsx'
            )
            
    except Exception as e:
        return f"Conversion error: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)