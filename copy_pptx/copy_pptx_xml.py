import logging
from io import StringIO

from pptx import Presentation
from pathlib import Path
import warnings
import zipfile
import xml.etree.ElementTree as ET
import os
import glob
import re
from pprint import pprint
from lxml import etree
import shutil

script_location = Path(__file__).absolute().parent


class CopyPptx:
    source_folder = f"{script_location}/source_pptx_extracted"
    target_indexes = dict()

    def __init__(self, path_to_source, path_to_new, slides_to_copy):
        self.path_to_source = path_to_source
        self.path_to_new = path_to_new
        self.slides_to_copy = slides_to_copy

    def copy_slides(self):
        os.makedirs(self.source_folder, exist_ok=True)

        with zipfile.ZipFile(self.path_to_source, 'r') as source_zip:
            source_zip.extractall(self.source_folder)

        with (zipfile.ZipFile(self.path_to_new, "w") as target_zip):
            self._working_with_xml()

            self._copy_all_files(source_zip, target_zip)

        # shutil.rmtree(source_folder)

    def _working_with_xml(self):
        i = 0
        slides_path = f"{self.source_folder}/ppt/slides"
        temp_slides_path = f"{self.source_folder}/ppt/slides/temp"

        slides_path_rels = f"{self.source_folder}/ppt/slides/_rels"
        temp_slides_path_rels = f"{self.source_folder}/ppt/slides/_rels/temp"

        for slide in self.slides_to_copy:
            i += 1
            self.change_file_index(f"{slides_path}/slide{slide}.xml", f"{temp_slides_path}/slide{i}.xml")
            self.change_rels_file(f"{slides_path_rels}/slide{slide}.xml.rels",
                                  f"{temp_slides_path_rels}/slide{i}.xml.rels", i)

    def change_file_index(self, slides_path, slide_xml_path_new):
        temp_slides_path = slides_path.rsplit('/', 1)[0] + "/temp"
        tree = etree.parse(slides_path)
        if not os.path.exists(temp_slides_path):
            os.makedirs(temp_slides_path)
        tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_rels_file(self, slides_path, slide_xml_path_new, new_num):
        self.change_file_index(slides_path, slide_xml_path_new)
        tree = etree.parse(slide_xml_path_new)
        root = tree.getroot()

        namespaces = CopyPptx.get_name_spaces_by_filepath(slide_xml_path_new)

        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        self._deep_change_target_links_rels(relationship_elements)
        for rel in relationship_elements:
            pattern = r'../slides/slide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)

            if r_num != new_num \
                    and r_num is not None:
                rel.set('Target', f'../slides/slide{new_num}.xml')

        for rel in relationship_elements:
            pattern = r'../notesSlides/notesSlide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)
            if r_num != new_num \
                    and r_num is not None:
                rel.set('Target', f'../notesSlides/notesSlide{new_num}.xml')

        tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _deep_change_target_links_rels(self, relationship_elements):
        for rel in relationship_elements:
            target = str(rel.get('Target'))
            target_type = str(rel.get('Type')).split('/')[-1]
            path_to_lib = target.replace('..', self.source_folder)
            print(path_to_lib)
            index = self.add_target_indexes(target_type)

            # if target_type == 'chart':
            #     self.change_file_index(slides_path, re.sub(r'\d', index, original_string))


            # rel.set('Target', value)

    def _change_target_links_rels(self, relationship_elements, pattern, new_num, value):
        for rel in relationship_elements:
            r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)
            if r_num != new_num \
                    and r_num is not None:
                rel.set('Target', value)

    def _copy_all_files(self, source_zip, target_zip):
        """
        Adding common files for pptx from source_zip pptx
        :param source_zip: source pptx opened like zip
        :param target_zip: target pptx opened like zip
        :return: none
        """
        for file in source_zip.namelist():
            try:
                target_zip.write(os.path.join(self.source_folder, file), file)
            except OSError as e:
                logging.warning(e)

    @staticmethod
    def extract_slide_numbers(text, pattern):
        matches = re.findall(pattern, text)
        slide_numbers = [int(match) for match in matches]
        if len(slide_numbers) == 0:
            return None

        return slide_numbers[0]

    def add_target_indexes(self, target_type):
        if target_type in self.target_indexes:
            self.target_indexes[target_type] = self.target_indexes.get(target_type) + 1
        else:
            self.target_indexes[target_type] = 1

        return self.target_indexes[target_type]

    @staticmethod
    def get_name_spaces(root):
        return dict([
            node for _, node in ET.iterparse(StringIO(str(ET.tostring(root), encoding='utf-8')), events=['start-ns'])])

    @staticmethod
    def get_name_spaces_by_filepath(filepath):
        return dict([node for _, node in ET.iterparse(filepath,
                                                      events=['start-ns'])])

    @staticmethod
    def delete_child(rel):
        parent = rel.getparent()
        if parent is not None:
            parent.remove(rel)


def main():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)
    path_to_source = f"{script_location}/template.pptx"

    pptx_copy = CopyPptx(path_to_source, path_to_new,
                         [2])

    pptx_copy.copy_slides()


main()

# [22, 23, 22, 23, 26, 26, 12,
#  12, 16, 17, 22, 23, 18, 16, 17, 22, 23,
#  16, 17, 16, 17, 32, 16, 17, 18, 18, 22,
#  23, 26, 26, 18, 16, 17, 18, 9, 16, 17,
#  16, 17, 22, 23, 16, 17, 18, 16, 17, 9,
#  16, 17, 18, 22, 23, 18, 9, 9, 18, 16,
#  17, 22, 23, 16, 17, 26, 16, 17, 18, 16,
#  17, 18, 16, 17, 16, 17, 16, 17, 18, 18, 16,
#  17, 16, 17, 18, 16, 17, 16, 17, 16, 17, 26]
