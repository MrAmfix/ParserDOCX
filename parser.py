import json
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from schemas import schemas
from typing import List, Optional, Dict
from os.path import join, dirname, abspath


class Elem:
    def __init__(self, num: str, text: str):
        self.num: str = num
        self.text: str = text
        self.sub_elements: Optional[List[Elem]] = None

    def append(self, elem):
        if self.sub_elements is None:
            self.sub_elements = []
        self.sub_elements.append(elem)

    def __len__(self):
        return 0 if self.sub_elements is None else len(self.sub_elements)


class Parser:
    def __init__(self, path: str):
        self.path = path
        self._temp_dir = tempfile.mkdtemp()
        self._extract_files()

    def _extract_files(self):
        with zipfile.ZipFile(self.path, 'r') as zr:
            zr.extractall(self._temp_dir)

    def parse(self):
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        paragraphs = []
        for element in root.iter(f'{{{schemas.w}}}sdt'):
            for content in element.iter(f'{{{schemas.w}}}sdtContent'):
                for para in content.iter(f'{{{schemas.w}}}p'):
                    paragraphs.append(para)
        return self._parse_paragraphs(paragraphs)

    def _parse_paragraphs(self, paragraphs: List[Element]):
        struct = []
        potentially_damage = False
        for para in paragraphs:
            for pPr in para.iter(f'{{{schemas.w}}}pPr'):
                for style in pPr.iter(f'{{{schemas.w}}}pStyle'):
                    val = style.attrib[f'{{{schemas.w}}}val']
                    if not val.isdigit():
                        continue
                    elif int(val) // 10 == 1:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            struct.append(Elem(str(len(struct) + 1), text.text))
                            break
                    elif int(val) // 10 == 2:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            if len(struct) == 0:
                                struct.append(Elem(str(len(struct) + 1), text.text))
                                potentially_damage = True
                                break
                            else:
                                num = f'{str(len(struct))}.{str(len(struct[-1]) + 1)}'
                                struct[-1].append(Elem(num, text.text))
                                break
                    elif int(val) // 10 == 3:
                        for text in para.iter(f'{{{schemas.w}}}t'):
                            if len(struct) == 0:
                                potentially_damage = True
                                struct.append(Elem(str(len(struct) + 1), text.text))
                                break
                            elif len(struct[-1]) == 0:
                                potentially_damage = True
                                num = f'{str(len(struct))}.{str(len(struct[-1]) + 1)}'
                                struct[-1].append(Elem(num, text.text))
                                break
                            else:
                                num = (f'{str(len(struct))}.{str(len(struct[-1]))}.'
                                       f'{str(len(struct[-1].sub_elements) + 1)}')
                                struct[-1].sub_elements[-1].append(Elem(num, text.text))
                                break
        return struct, potentially_damage

    def save(self, struct: Dict, path: str = None):
        if path is None:
            path = self.path
        with open(f'{path[:-5]}.json', 'w', encoding='utf-8') as json_file:
            json.dump(struct, json_file, indent=4, default=handle_none,
                      ensure_ascii=False)

    def get_other_text(self):
        tree = ET.parse(os.path.join(self._temp_dir, 'word', 'document.xml'))
        root = tree.getroot()
        all_text = ''
        for para in root.findall(f'.//{{{schemas.w}}}body//{{{schemas.w}}}p'):
            for t in para.iter(f'{{{schemas.w}}}t'):
                all_text += f'{t.text}\n'
        return all_text


def struct_to_dict(elements: List[Elem]):
    if elements is None:
        return

    def convert_element_to_dict(elem: Elem):
        elem_dict = {
            "num": elem.num,
            "text": elem.text,
        }
        if elem.sub_elements:
            elem_dict["sub_elements"] = [convert_element_to_dict(sub_elem) for sub_elem in elem.sub_elements]
        return elem_dict
    return [convert_element_to_dict(elem) for elem in elements]


def handle_none(obj):
    if obj is None:
        return "null"
    return obj


def iter_docs():
    root = str(join(dirname(abspath(__file__))))
    for filename in os.listdir(join(root, 'documents_for_extract')):
        print(filename)


def parsing_documents():
    root_in = join(str(join(dirname(abspath(__file__)))), 'documents_for_extract')
    root_out = join(str(join(dirname(abspath(__file__)))), 'out_json')
    for filename in os.listdir(root_in):
        try:
            p = Parser(join(root_in, filename))
            s, pot = p.parse()
            n = {'potentially_damage': pot, 'table_of_content': struct_to_dict(s),
                 'other_text': p.get_other_text()}
            p.save(n, join(root_out, filename))
        except Exception as _e:
            continue


if __name__ == '__main__':
    parsing_documents()
