import zipfile
import xml.etree.ElementTree as ET

z = zipfile.ZipFile('data/04_prompt/test_1x1.xlsx')
xml_bytes = z.read('xl/worksheets/sheet1.xml')
root = ET.fromstring(xml_bytes)
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

sheet_pr = root.find('s:sheetPr', ns)
if sheet_pr is not None:
    page_setup_pr = sheet_pr.find('s:pageSetUpPr', ns)
    if page_setup_pr is not None:
        print('pageSetUpPr attributes:', page_setup_pr.attrib)
    else:
        print('pageSetUpPr not found')
else:
    print('sheetPr not found')

page_setup = root.find('s:pageSetup', ns)
if page_setup is not None:
    print('pageSetup attributes:', page_setup.attrib)
else:
    print('pageSetup not found')
