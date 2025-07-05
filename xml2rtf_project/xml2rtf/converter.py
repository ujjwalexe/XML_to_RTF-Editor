from .parser import parse_xml
from .formatter import format_rtf
from .utils import init_rtf_header, close_rtf

def convert(xml_string: str) -> str:
    paragraphs = parse_xml(xml_string)
    rtf_content = init_rtf_header()

    for para in paragraphs:
        rtf_content += format_rtf(para)

    rtf_content += close_rtf()
    return rtf_content
