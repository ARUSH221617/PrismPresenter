import sys
import zipfile
import lxml.etree as etree

def clean_element(elem):
    # If element is an a:p or a:r or similar container, ensure elem.text and child tails are None
    tag = etree.QName(elem.tag).localname if elem.tag else ""
    if tag in ("p", "r", "pPr", "rPr", "bodyPr", "txBody", "spPr", "nvSpPr"):
        elem.text = None
    for child in elem:
        clean_element(child)
        child.tail = None

def fix_file_lxml(path: str):
    temp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                    root = etree.fromstring(raw)
                    clean_element(root)
                    raw = etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes")
                zout.writestr(item, raw)
    import os
    os.replace(temp, path)

fix_file_lxml("data/output/t8.pptx")
print("Cleaned t8.pptx with lxml")

sys.path.insert(0, "src")
from pptx_jahat.tools.renderers.com_renderer import export_pptx_slides_com
imgs = export_pptx_slides_com("data/output/t8.pptx", width=750)
print("SUCCESS EXPORTED:", len(imgs))
