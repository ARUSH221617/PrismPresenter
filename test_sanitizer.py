import sys
import zipfile
import re
import os

def sanitize_slide_xml(xml_bytes: bytes) -> bytes:
    """
    Cleans mixed-content text that was accidentally placed directly under <a:p> or <a:r>.
    In OpenXML DrawingML, <a:p> must only contain elements (<a:pPr>, <a:r>, <a:endParaRPr>, etc.)
    and <a:r> must only contain <a:rPr> and <a:t>.
    Direct text nodes between tags corrupt PowerPoint's native Slide.Export XML reader.
    """
    # Regex pattern matching text immediately after <a:p> or <a:r> before child tag
    # Example: <a:p>some text<a:pPr> -> <a:p><a:pPr>
    # Example: <a:r>some text<a:rPr> -> <a:r><a:rPr>
    cleaned = re.sub(rb'(<a:p(?: [^>]*)?>)([^<]+)(<a:pPr|<a:r|<a:endParaRPr)', rb'\1\3', xml_bytes)
    cleaned = re.sub(rb'(<a:r(?: [^>]*)?>)([^<]+)(<a:rPr|<a:t)', rb'\1\3', cleaned)
    return cleaned

def clean_pptx_file(pptx_path: str):
    p = os.path.abspath(pptx_path)
    temp_p = p + ".tmp"
    with zipfile.ZipFile(p, "r") as zin:
        with zipfile.ZipFile(temp_p, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                    data = sanitize_slide_xml(data)
                zout.writestr(item, data)
    os.replace(temp_p, p)

clean_pptx_file("data/output/t8.pptx")
print("Sanitized data/output/t8.pptx")

sys.path.insert(0, "src")
from pptx_jahat.tools.renderers.com_renderer import export_pptx_slides_com
imgs = export_pptx_slides_com("data/output/t8.pptx", width=750)
print("SUCCESSFULLY EXPORTED T8.PPTX:", len(imgs), "slides via Native PowerPoint COM!")
