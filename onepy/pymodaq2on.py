import os, sys
import pytz
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.dirname(SCRIPT_DIR))

from onepy import OneNote
from onmanager import ONProcess


def generateOneNoteXML(page_id, bold_text, normal_text_list, style='\"font-family:Arial;font-size:15.0pt\"'):
    xml = '<Page ID=\"' + str(page_id) + '\" xmlns=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">' + \
          '<Outline>' + \
          '<OEChildren>' + \
          '<OE style=' + style + '>'

    xml += '<T><![CDATA[<b>' + bold_text + '</b>]]></T>'

    for item in normal_text_list:
        xml += '<T><![CDATA[' + item + ']]></T>'

    xml += '</OE>' + \
           '</OEChildren>' + \
           '</Outline>' + \
           '</Page>'

    return xml


def writeToOneNote(notebook_name, bold_text, normal_text_list, style='\"font-family:Arial;font-size:15.0pt\"'):
    on = OneNote()

    # Look for the last page
    for nbk in on.hierarchy:
        if nbk.name == notebook_name:
            for section in nbk:
                page = section._children[-1]
                print("Page = " + str(page) + " ID = " + str(page.id))

    # Stupid hack
    date = pytz.utc.localize(datetime(year=1899, month=12, day=30))

    xml = generateOneNoteXML(page.id, bold_text, normal_text_list, style)
    on.process.update_page_content(xml, date)


def main():
    notebook_name = "test_python"
    bold_text = 'Scan'
    normal_text_list = ['One\n', 'Two', 'Three']
    writeToOneNote(notebook_name, bold_text, normal_text_list)


if __name__ == "__main__":
    main()
