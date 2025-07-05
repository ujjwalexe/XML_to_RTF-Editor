# xml2rtf/parser.py

from lxml import etree

def parse_xml(xml_string):
    root = etree.fromstring(xml_string.encode('utf-8'))
    paragraphs = []

    for p in root.findall(".//p"):
        style = p.attrib.get("style", "")
        runs = []

        for r in p.findall("r"):
            attribs = r.attrib
            text_node = r.find("t")
            text = text_node.text if text_node is not None else ""
            runs.append({"text": text, "attrib": attribs})

        paragraphs.append({"style": style, "runs": runs})
    
    return paragraphs
