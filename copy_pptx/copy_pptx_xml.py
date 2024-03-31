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
        slides_path_note = f"{self.source_folder}/ppt/notesSlides"
        slides_path_charts = f"{self.source_folder}/ppt/charts"
        slides_path_embeddings = f"{self.source_folder}/ppt/embeddings"
        for slide in self.slides_to_copy:
            i += 1
            self._change_file_index(f"{slides_path}/slide{slide}.xml", i)
            self._change_rels_file(f"{slides_path}/slide{slide}.xml", i)

        CopyPptx.delete_files_from_folder(slides_path, 'slide*')
        CopyPptx.delete_files_from_folder(slides_path_note, 'notesSlide*')
        CopyPptx.delete_files_from_folder(slides_path_charts, 'charts*')
        CopyPptx.delete_files_from_folder(slides_path_embeddings, 'Microsoft_Excel_Worksheet*')

        CopyPptx.move_all_files(slides_path)
        CopyPptx.move_all_files(slides_path_note)
        CopyPptx.move_all_files(slides_path_charts)
        CopyPptx.move_all_files(slides_path_embeddings)

    def _change_rels_file(self, slides_path, new_index):
        slide_xml_path_new = self._change_file_index_rels(slides_path, new_index)
        relationship_elements = CopyPptx._get_relationship_elements(slide_xml_path_new)
        self._deep_change_target_links_rels(relationship_elements)
        for rel in relationship_elements:
            pattern = r'../slides/slide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)

            if r_num != new_index \
                    and r_num is not None:
                rel.set('Target', f'../slides/slide{new_index}.xml')

        for rel in relationship_elements:
            pattern = r'../notesSlides/notesSlide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)
            if r_num != new_index \
                    and r_num is not None:
                rel.set('Target', f'../notesSlides/notesSlide{new_index}.xml')

        # tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    # TODO:: не забыть переместить и удалить файлы
    def _deep_change_target_links_rels(self, relationship_elements):
        for rel in relationship_elements:
            target = str(rel.get('Target'))
            target_type = str(rel.get('Type')).split('/')[-1]
            path_to_lib = target.replace('..', self.source_folder + '/ppt')

            index = self.add_target_indexes(target_type)

            if target_type == 'chart':
                self._change_chart(path_to_lib, index)
                rel.set('Target', f'../charts/charts{index}.xml')

            if target_type == 'notesSlides':
                pattern = r'../notesSlides/notesSlide(\d+)\.xml'
                r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)
                if r_num != index \
                        and r_num is not None:
                    rel.set('Target', f'../notesSlides/notesSlide{index}.xml')
                self._change_notes_slides(path_to_lib, index)
        # rel.set('Target', value)

    def _change_notes_slides(self, path_to_lib, index):
        CopyPptx._change_file_index(path_to_lib, index)
        notes_slides_rels = CopyPptx._change_file_index_rels(path_to_lib, index)

        relationship_elements = CopyPptx._get_relationship_elements(notes_slides_rels)
        for notes_rel in relationship_elements:
            pattern = r'../slides/slide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(notes_rel.get('Target'), pattern)

            if r_num != index \
                    and r_num is not None:
                notes_rel.set('Target', f'../slides/slide{index}.xml')

    def _change_chart(self, path_to_lib, index):
        CopyPptx._change_file_index(path_to_lib, index)
        slide_xml_path_new = CopyPptx._change_file_index_rels(path_to_lib, index)
        relationship_elements = CopyPptx._get_relationship_elements(slide_xml_path_new)
        for rel in relationship_elements:
            chart_target = str(rel.get('Target'))
            chart_target_type = str(rel.get('Type')).split('/')[-1]
            embedding_index = self.add_target_indexes(chart_target_type)

            chart_path_to_lib = chart_target.replace('..', self.source_folder + '/ppt')
            if embedding_index == 1:
                CopyPptx._change_file_index(chart_path_to_lib, "")
            else:
                CopyPptx._change_file_index(chart_path_to_lib, embedding_index)
                rel.set('Target', CopyPptx.replace_number(chart_target, embedding_index))

    @staticmethod
    def _get_relationship_elements(slide_xml_path):
        tree = etree.parse(slide_xml_path)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(slide_xml_path)
        return root.findall('.//Relationship', namespaces=namespaces)

    @staticmethod
    def _change_target_links_rels(relationship_elements, pattern, new_num, value):
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

    def add_target_indexes(self, target_type):
        if target_type in self.target_indexes:
            self.target_indexes[target_type] = self.target_indexes.get(target_type) + 1
        else:
            self.target_indexes[target_type] = 1

        return self.target_indexes[target_type]

    @staticmethod
    def delete_files_from_folder(slides_path, pattern):
        CopyPptx.delete_all_files(slides_path, pattern)
        CopyPptx.delete_all_files(slides_path + '/_rels', pattern)

    @staticmethod
    def delete_all_files(slides_path, pattern):
        files_to_delete = glob.glob(os.path.join(slides_path, pattern))
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
            except OSError as e:
                logging.warning(f"Error of deleting file '{file_path}': {e}")

    @staticmethod
    def extract_slide_numbers(text, pattern):
        matches = re.findall(pattern, text)
        slide_numbers = [int(match) for match in matches]
        if len(slide_numbers) == 0:
            return None

        return slide_numbers[0]

    @staticmethod
    def _change_file_index_rels(slides_path, new_index):
        temp_slides_path = slides_path.rsplit('/', 1)[0] + '/_rels' + '/temp'
        CopyPptx.create_a_dir(temp_slides_path)
        tree = etree.parse(slides_path)
        new_file_number = re.sub(r'\d', str(new_index), str(slides_path.rsplit('/', 1)[1])) + '.rels'
        result_path = f"{temp_slides_path}/{new_file_number}"
        tree.write(result_path, pretty_print=True, xml_declaration=True, encoding='utf-8')
        return result_path

    @staticmethod
    def _change_file_index(slides_path, new_index):
        temp_slides_path = slides_path.rsplit('/', 1)[0] + "/temp"
        tree = etree.parse(slides_path)
        CopyPptx.create_a_dir(temp_slides_path)
        new_file_number = re.sub(r'\d', str(new_index), str(slides_path.rsplit('/', 1)[1]))
        tree.write(temp_slides_path + new_file_number, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def replace_number(source_str, new_index):
        re.sub(r'\d', str(new_index), str(source_str))

    @staticmethod
    def create_a_dir(path):
        if not os.path.exists(path):
            os.makedirs(path)

    @staticmethod
    def get_name_spaces(root):
        return dict([
            node for _, node in ET.iterparse(StringIO(str(ET.tostring(root), encoding='utf-8')), events=['start-ns'])])

    @staticmethod
    def move_all_files(source_folder):
        CopyPptx.move_files(source_folder)
        CopyPptx.move_files(source_folder + '/_rels')

    @staticmethod
    def move_files(source_folder):
        temp_folder = source_folder + "/temp"
        if not os.path.exists(temp_folder):
            logging.warning(f"Source folder '{source_folder}' can not be find.")
        else:
            for filename in os.listdir(temp_folder):
                file_path = os.path.join(temp_folder, filename)
                if os.path.isfile(file_path):
                    shutil.move(file_path, source_folder)

            shutil.rmtree(temp_folder)

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
