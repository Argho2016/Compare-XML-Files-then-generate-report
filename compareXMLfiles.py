import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font

def parse_xml(file_path):
    tree = ET.parse(file_path)
    return tree.getroot()

def compare_elements(elem1, elem2, path=''):
    diffs = []

    if elem1.tag != elem2.tag:
        diffs.append(f"{path}: Tag differs: '{elem1.tag}' vs '{elem2.tag}'")
    if elem1.text != elem2.text:
        diffs.append(f"{path}: Text differs: '{elem1.text}' vs '{elem2.text}'")

    for attr in set(elem1.attrib) | set(elem2.attrib):
        val1 = elem1.attrib.get(attr)
        val2 = elem2.attrib.get(attr)
        if val1 != val2:
            diffs.append(f"{path}: Attribute '{attr}' differs: '{val1}' vs '{val2}'")

    children1 = list(elem1)
    children2 = list(elem2)
    for i, (child1, child2) in enumerate(zip(children1, children2)):
        child_path = f"{path}/{elem1.tag}[{i}]"
        diffs.extend(compare_elements(child1, child2, child_path))
    
    if len(children1) > len(children2):
        for i in range(len(children2), len(children1)):
            diffs.append(f"{path}/{elem1.tag}[{i}]: Missing in second file")
    elif len(children2) > len(children1):
        for i in range(len(children1), len(children2)):
            diffs.append(f"{path}/{elem2.tag}[{i}]: Missing in first file")
    
    return diffs

def generate_report(diffs, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "XML Comparison Report"

    header_font = Font(bold=True)
    sheet["A1"] = "Differences"
    sheet["A1"].font = header_font

    for idx, diff in enumerate(diffs, start=2):
        sheet[f"A{idx}"] = diff

    workbook.save(output_file)

def main(file1, file2, output_file):
    root1 = parse_xml(file1)
    root2 = parse_xml(file2)

    diffs = compare_elements(root1, root2)

    generate_report(diffs, output_file)

if __name__ == "__main__":
    file1 = "36.xml"
    file2 = "96.xml"
    output_file = "xml_comparison_report.xlsx"

    main(file1, file2, output_file)
