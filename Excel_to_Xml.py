import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re

def clean_tag_name(tag_name):
    # Replace invalid XML tag characters with underscores
    tag_name = re.sub(r'\s+', '_', tag_name)
    tag_name = re.sub(r'[^a-zA-Z0-9_]', '', tag_name)
    return tag_name

def excel_to_xml(excel_file, xml_file):
    # Read Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Create root element
    root = ET.Element("data")

    # Iterate through rows
    for _, row in df.iterrows():
        # Create element for each row
        row_element = ET.SubElement(root, "row")

        # Iterate through columns
        for col_name, col_value in row.items():
            # Clean column name to make a valid XML tag
            col_name = clean_tag_name(col_name)

            # Create element for each column
            col_element = ET.SubElement(row_element, col_name)
            col_element.text = str(col_value)

    # Create and save XML file
    xml_str = ET.tostring(root, encoding="utf-8")
    xml_pretty_str = minidom.parseString(xml_str).toprettyxml(indent="  ")

    with open(xml_file, "w", encoding="utf-8") as xml_output:
        xml_output.write(xml_pretty_str)

# Example usage
excel_file_path = "/home/karam/Desktop/delai1.xlsx"  # Replace with your Excel file path
xml_file_path = "/home/karam/Desktop/converted1.xml"  # Replace with your desired output XML file path

excel_to_xml(excel_file_path, xml_file_path)