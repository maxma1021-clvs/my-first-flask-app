import xml.etree.ElementTree as ET
from typing import Dict, Any, List
from .utils import NS, parse_wordart

class VMLParser:
    def parse(self, root: ET.Element, part_name: str) -> List[Dict[str, Any]]:
        results = []
        for pict in root.findall(".//w:pict", NS):
            wordart_info = parse_wordart(pict, part_name, {})
            if wordart_info:
                results.append(wordart_info)
        return results
