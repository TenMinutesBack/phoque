import xml.etree.ElementTree as ET
import os

def action_examiner(case_dir):
    root = ET.parse(case_dir).getroot()
    # if 'action' in [elem.tag for elem in root.iter()]:
    #     return True
    for elem in root.iter():
        if elem.tag == 'operation':
            if elem.attrib['path'].split('/')[-1] == ''



action_examiner('tc_gb_usb_update_0080.xml')