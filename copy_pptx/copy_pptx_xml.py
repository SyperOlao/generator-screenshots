import logging
from io import StringIO
import random

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


def generate_ids(n):
    number = 3900
    counter = 0
    step = 3
    numbers = []
    for _ in range(n):
        numbers.append(number)
        counter += 1
        if counter == 6:
            step = 1
            counter = 0
        number += step
    return numbers


class CopyPptx:
    source_folder = f"{script_location}/source_pptx_extracted"
    target_indexes = dict()
    repeated_indexes = dict()
    styles = []

    def __init__(self, path_to_source, path_to_new, slides_to_copy):
        self.path_to_source = path_to_source
        self.path_to_new = path_to_new
        self.slides_to_copy = slides_to_copy
        self.len_master_id = 0
        self.get_repeated_indexes(slides_to_copy)

    def copy_slides(self):
        shutil.rmtree(self.source_folder)
        os.makedirs(self.source_folder, exist_ok=True)

        with zipfile.ZipFile(self.path_to_source, 'r') as source_zip:
            source_zip.extractall(self.source_folder)

        with (zipfile.ZipFile(self.path_to_new, "w") as target_zip):
            self._working_with_xml()

            self._copy_all_files(target_zip)

        # shutil.rmtree(self.source_folder)

    def _working_with_xml(self):

        slides_path = f"{self.source_folder}/ppt/slides"
        self.change_root_pptx_xml()
        self.change_root_pptx_xml_rels()
        self.change_doc_props()
        i = 0
        for slide in self.slides_to_copy:
            i += 1
            self.change_slide_id(slides_path + '/slide', slide)
            self._change_file_index(f"{slides_path}/slide{slide}.xml", i)
            self._change_rels_file(f"{slides_path}/slide{slide}.xml", i, slide)

        self.change_root_context_type()
        self.delete_and_move_files(slides_path)

    def change_slide_id(self, slides_path, old_index):
        slide_xml_path = f"{slides_path}{old_index}.xml"
        tree = etree.parse(slide_xml_path)
        root = tree.getroot()
        namespaces = self.get_name_spaces_by_filepath(slide_xml_path)
        creation_id = root.find('.//p14:creationId', namespaces=namespaces)
        if creation_id is not None:
            creation_id.set('val', f'{random.sample(range(2000000000, 7000000000), 1)[0]}')

        tree.write(slide_xml_path, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_doc_props(self):
        root_pptx_xml = f"{self.source_folder}/docProps/app.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = self.get_name_spaces_by_filepath(root_pptx_xml)

        slides = root.find('.//Slides', namespaces=namespaces)
        notes = root.find('.//Notes', namespaces=namespaces)
        current_slides_count = len(self.slides_to_copy)
        slides.text = str(current_slides_count)
        notes.text = str(current_slides_count)
        relationship_elements = root.findall('.//vt:lpstr', namespaces=namespaces)

        parent = None

        for rel in relationship_elements:
            if "PowerPoint" in rel.text:
                parent = rel.getparent()
                self.delete_child(rel)
        for i in range(len(self.slides_to_copy)):
            a = etree.SubElement(parent, "{" + namespaces['vt'] + "}lpstr")
            a.text = "Презентация PowerPoint"
        tree.write(root_pptx_xml)

    def change_root_context_type(self):
        root_pptx_xml = f"{self.source_folder}/[Content_Types].xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = self.get_name_spaces_by_filepath(root_pptx_xml)
        content_type = dict()
        relationship_elements = root.findall('.//Override', namespaces=namespaces)
        relations = None
        for rel in relationship_elements:
            target_type = str(rel.get('ContentType')).split('.')[-1].split('+')[0]
            content_type[target_type] = {'ct': str(rel.get('ContentType')), 'pt': str(rel.get('PartName'))}
            if target_type == 'slide' or target_type == 'chart' or target_type == 'notesSlide':
                relations = rel.getparent()
                CopyPptx.delete_child(rel)

        for elem in self.styles:
            type = CopyPptx.extract_before_first_number(elem)
            ct = content_type["chartstyle"]['ct']

            if type == 'colors':
                ct = content_type["chartcolorstyle"]['ct']

            etree.SubElement(relations, "Override",
                             {
                                 "PartName": f"/ppt/charts/{elem}",
                                 "ContentType": ct
                             })

        for i in range(len(self.slides_to_copy)):
            etree.SubElement(relations, "Override",
                             {
                                 "PartName": CopyPptx.replace_number(f"{content_type['slide']['pt']}",
                                                                     str(i + 1)),
                                 "ContentType": f"{content_type['slide']['ct']}"
                             })
        for target_type in self.target_indexes:
            if target_type not in content_type:
                continue
            for i in range(self.target_indexes[target_type]):
                etree.SubElement(relations, "Override",
                                 {
                                     "PartName": CopyPptx.replace_number(f"{content_type[target_type]['pt']}",
                                                                         str(i + 1)),
                                     "ContentType": f"{content_type[target_type]['ct']}"
                                 })

        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def extract_before_first_number(s):
        match = re.search(r'^([^\d]*)', s)
        if match:
            return match.group(1)
        else:
            return s

    def change_root_pptx_xml_rels(self):
        root_pptx_xml = f"{self.source_folder}/ppt/_rels/presentation.xml.rels"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(root_pptx_xml)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        all_type = ""
        relations = None
        for rel in relationship_elements:
            target_type = str(rel.get('Type')).split('/')[-1]
            if target_type == 'slide':
                all_type = str(rel.get('Type'))
                relations = rel.getparent()
                if relations is not None:
                    relations.remove(rel)

        for i in range(len(self.slides_to_copy)):
            etree.SubElement(relations, "Relationship",
                             {'Id': f'rId{self.len_master_id + i + 1}',
                              "Type": f'{all_type}',
                              "Target": f'slides/slide{i + 1}.xml'})
        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_root_pptx_xml(self):
        root_pptx_xml = f"{self.source_folder}/ppt/presentation.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = self.get_name_spaces(root)
        sld_ids = root.xpath('//ns0:sldId', namespaces=namespaces)
        self.len_master_id = len(root.xpath('//ns0:sldMasterId', namespaces=namespaces))

        for sldId in sld_ids:
            CopyPptx.delete_child(sldId)
        sldIdLst = root.find('ns0:sldIdLst', namespaces=namespaces)
        ids = generate_ids(len(self.slides_to_copy))
        for i in range(len(self.slides_to_copy)):
            etree.SubElement(sldIdLst, "{" + namespaces['ns0'] + "}sldId",
                             {'id': f'{str(ids[i])}',
                              "{" + namespaces['ns1'] + "}id": f'rId{i + 1 + self.len_master_id}'})

        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_rels_file(self, slides_path, new_index, old_index):
        slide_xml_path_new = self._change_file_index_rels(slides_path, new_index)
        self._deep_change_target_links_rels(slide_xml_path_new, old_index)

    def _deep_change_target_links_rels(self, slide_xml_path_new, old_index):
        tree = etree.parse(slide_xml_path_new)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(slide_xml_path_new)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        notes_slides_path = f"{self.source_folder}/ppt/notesSlides/notesSlide"
        if old_index in self.repeated_indexes:
            self.repeated_indexes[old_index] += 1

        for rel in relationship_elements:
            target = str(rel.get('Target'))
            target_type = str(rel.get('Type')).split('/')[-1]
            path_to_rel = target.replace('..', self.source_folder + '/ppt')
            index = self.add_target_indexes(target_type)
            if target_type == 'chart':
                self._change_chart_id(path_to_rel)
                self._change_chart_rels(path_to_rel, index, old_index)
                rel.set('Target', f'../charts/chart{index}.xml')

            if target_type == 'notesSlide':
                pattern = r'../notesSlides/notesSlide(\d+)\.xml'
                r_num = CopyPptx.extract_slide_numbers(rel.get('Target'), pattern)
                if r_num != index \
                        and r_num is not None:
                    rel.set('Target', f'../notesSlides/notesSlide{index}.xml')
                self.change_slide_id(notes_slides_path, old_index)
                self._change_notes_slides(path_to_rel, index)
        tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def _change_notes_slides(path_to_lib, index):
        CopyPptx._change_file_index(path_to_lib, index)
        notes_slides_rels = CopyPptx._change_file_index_rels(path_to_lib, index)
        tree = etree.parse(notes_slides_rels)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(notes_slides_rels)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        for notes_rel in relationship_elements:
            pattern = r'../slides/slide(\d+)\.xml'
            r_num = CopyPptx.extract_slide_numbers(notes_rel.get('Target'), pattern)
            if r_num != index \
                    and r_num is not None:
                notes_rel.set('Target', f'../slides/slide{index}.xml')
        tree.write(notes_slides_rels, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def generate_hex_string():
        group1 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group2 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group3 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group4 = ''.join(random.choices('0123456789ABCDEF', k=12))
        formatted_hex_string = f"{group1}-{group2}-{group3}-{group4}"
        return formatted_hex_string

    @staticmethod
    def _change_chart_id(path_to_chart):
        tree = etree.parse(path_to_chart)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(path_to_chart)
        # TODO: THINK ABOUT ID
        try:
            unique_ids = root.findall('.//c16:uniqueId', namespaces=namespaces)
            hex = CopyPptx.generate_hex_string()
            # {00000003-1A3C-46CD-A049-AE5FE0497CE7}
            # {00000001-EB99-668B-87F1-6F6B
            # 920M-B6CJ-OC8Y-0B0NXSWHZRJV
            for id in unique_ids:
                curr_id = str(id.get('val')).split('-')[0][1::]

                id.set('val', "{" + f"{curr_id}-{hex}" + "}")
            # c16:uniqueId
            tree.write(path_to_chart, pretty_print=True, xml_declaration=True, encoding='utf-8')
        except Exception as err:
            logging.warning(err)

    def _change_chart_rels(self, path_to_chart, index, old_index):
        CopyPptx._change_file_index(path_to_chart, index)
        chart_path_rels = CopyPptx._change_file_index_rels(path_to_chart, index)

        tree = etree.parse(chart_path_rels)
        root = tree.getroot()
        namespaces = CopyPptx.get_name_spaces_by_filepath(chart_path_rels)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)

        for rel in relationship_elements:
            chart_target = str(rel.get('Target'))
            chart_target_type = str(rel.get('Type')).split('/')[-1]
            if chart_target_type == 'package':
                self.change_package(rel, chart_target_type, chart_target)
            if chart_target_type == 'chartStyle':
                self.change_chart_style(rel, chart_target, 'style*', old_index)
            if chart_target_type == 'chartColorStyle':
                self.change_chart_style(rel, chart_target, 'colors*', old_index)

        tree.write(chart_path_rels, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_chart_style(self, rel, chart_target, pattern, old_index):

        if old_index not in self.repeated_indexes or self.repeated_indexes[old_index] < 2:
            return

        chart_path_to_embedding = self.source_folder + '/ppt/charts'
        index = CopyPptx.get_last_index(chart_path_to_embedding, pattern)

        new_chart_style = CopyPptx.replace_number(chart_target, str(index + 1))
        self.styles.append(new_chart_style)

        old_path = chart_path_to_embedding + '/' + chart_target
        new_path = chart_path_to_embedding + '/' + new_chart_style
        shutil.copy2(old_path, new_path)

        # print("old_path", old_path)
        # print("new_path", new_path)
        rel.set('Target', new_chart_style)

    def change_package(self, rel, chart_target_type, chart_target):
        embedding_index = str(self.add_target_indexes(chart_target_type))

        chart_path_to_embedding = chart_target.replace('..', self.source_folder + '/ppt')

        new_chart_path = os.path.dirname(chart_path_to_embedding) + '/temp'

        new_name = CopyPptx.get_embedding_name(chart_path_to_embedding, embedding_index)

        CopyPptx.rename_and_move_file(chart_path_to_embedding,
                                      new_name, new_chart_path)

        rel.set('Target', f'../embeddings/{new_name}')

    @staticmethod
    def get_last_index(path, pattern):
        files_to_index = glob.glob(os.path.join(path, pattern))
        num = []
        for file in files_to_index:
            num.append(int(CopyPptx.get_number_from_str(file)[0]))

        return max(num)

    @staticmethod
    def get_number_from_str(local_string):
        return re.findall(r'\d+', str(local_string))

    @staticmethod
    def get_embedding_name(chart_path_to_embedding, embedding_index):
        excel_name = chart_path_to_embedding.split('/')[-1]
        if embedding_index == '1':
            return re.sub(r'\d+', '', excel_name)

        if re.search(r'\d', excel_name):
            new_name = str(CopyPptx.replace_number(excel_name,
                                                   str(int(embedding_index) - 1)))
        else:
            new_name = excel_name.split('.')[0] + str(int(embedding_index) - 1) + '.xlsx'

        if new_name is None:
            return str(excel_name).join('')
        return new_name

    def _copy_all_files(self, target_zip):
        """
        Adding common files for pptx from source_folder
        :param source_folder: source folder path
        :param target_zip: target pptx opened like zip
        :return: none
        """
        for root, dirs, files in os.walk(self.source_folder):
            for file in files:
                source_path = os.path.join(root, file)
                relative_path = os.path.relpath(source_path, self.source_folder)
                try:
                    target_zip.write(source_path, relative_path)
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
        slides_path_2 = slides_path.rsplit('/', 1)[0] + '/_rels'
        file = slides_path_2 + '/' + slides_path.rsplit('/', 1)[1] + '.rels'
        temp_slides_path = slides_path_2 + '/temp'
        CopyPptx.create_a_dir(temp_slides_path)
        tree = etree.parse(file)
        new_file_number = re.sub(r'\d+', str(new_index), str(slides_path.rsplit('/', 1)[1])) + '.rels'
        result_path = f"{temp_slides_path}/{new_file_number}"
        tree.write(result_path, pretty_print=True, xml_declaration=True, encoding='utf-8')
        return result_path

    @staticmethod
    def _change_file_index(slides_path, new_index):
        temp_slides_path = slides_path.rsplit('/', 1)[0] + "/temp"
        tree = etree.parse(slides_path)
        CopyPptx.create_a_dir(temp_slides_path)
        new_file_number = re.sub(r'\d+', str(new_index), str(slides_path.rsplit('/', 1)[1]))
        tree.write(temp_slides_path + "/" + new_file_number, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def rename_and_move_file(old_path: str, new_name: str, new_directory):
        if not os.path.isfile(old_path):
            logging.warning(f"File {old_path} is not found.")
            return
        new_path = new_directory + "/" + new_name
        CopyPptx.create_a_dir(new_directory)

        shutil.copy2(old_path, new_path)
        logging.info(f"File has been renamed and moved: {old_path} -> {new_path}")

    @staticmethod
    def replace_number(source_str, new_index):
        if not re.search(r'\d', source_str):
            return source_str
        else:
            return re.sub(r'\d+', str(new_index), source_str)

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

    def get_repeated_indexes(self, numbers):
        counts = {}
        for num in numbers:
            if num in counts:
                counts[num] += 1
            else:
                counts[num] = 1
        rep = [num for num, count in counts.items() if count > 1]
        self.repeated_indexes = dict()
        for i in rep:
            self.repeated_indexes[i] = 0

    def delete_and_move_files(self, slides_path):
        slides_path_note = f"{self.source_folder}/ppt/notesSlides"
        slides_path_charts = f"{self.source_folder}/ppt/charts"
        slides_path_embeddings = f"{self.source_folder}/ppt/embeddings"

        CopyPptx.delete_files_from_folder(slides_path, 'slide*')
        CopyPptx.delete_files_from_folder(slides_path_note, 'notesSlide*')
        CopyPptx.delete_files_from_folder(slides_path_charts, 'chart*')
        CopyPptx.delete_files_from_folder(slides_path_embeddings, 'Microsoft_Excel_Worksheet*')

        CopyPptx.move_all_files(slides_path)
        CopyPptx.move_all_files(slides_path_note)
        CopyPptx.move_all_files(slides_path_charts)
        CopyPptx.move_files(slides_path_embeddings)


def search_word_in_xml_folder(folder_path, word):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xml') or file.endswith('.rels'):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r') as f:
                        line_number = 0
                        for line in f:
                            line_number += 1
                            if word in line:
                                print(f"Word '{word}' found in file: {file_path}, line {line_number}:")
                                print(line.strip())
                except Exception as e:
                    print(f"Error reading file {file_path}: {e}")


def search_word_in_xml_element(element, word):
    if element.text and word in element.text:
        return True
    for child in element:
        if search_word_in_xml_element(child, word):
            return True
    return False


def main():
    new_presentation = Presentation()
    path_to_temp = f"{script_location}/res_2.pptx"
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)
    path_to_source = f"{script_location}/template.pptx"

    # slides_to_copy = random.sample(range(1, 32), 31)
    slides_to_copy = [2, 2, 2, 28, 26, 30, 2, 9, 29, 14, 12, 15, 13, 6, 31, 27, 7, 28, 19]
    for i in range(len(slides_to_copy)):
        print("i: ", i + 1, " ", slides_to_copy[i])
    pptx_copy = CopyPptx(path_to_source, path_to_new,
                         slides_to_copy)
    sr = f"{script_location}/res_2"
    os.makedirs(sr, exist_ok=True)
    with zipfile.ZipFile(path_to_temp, 'r') as source_zip:
        source_zip.extractall(sr)

    folder_path = f"{script_location}/source_pptx_extracted"
    word_to_search = 'style2'
    search_word_in_xml_folder(folder_path, word_to_search)

    # [23, 17, 28, 8, 26, 30, 22, 19, 2, 21, 9, 29, 14, 12, 15, 13, 5, 24, 10, 25, 18, 4, 11, 16, 20, 1, 6, 31, 27, 7, 3]
    # pptx_copy = CopyPptx(path_to_source, path_to_new,
    #                      [22, 23, 22, 23, 26, 26, 12,
    #                       12, 16, 17, 22, 23, 18, 16, 17, 22, 23,
    #                       16, 17, 16, 17, 32, 16, 17, 18, 18, 22,
    #                       # 23, 26, 26, 18, 16, 17, 18, 9, 16, 17,
    #                       # 16, 17, 22, 23, 16, 17, 18, 16, 17, 9,
    #                       # 16, 17, 18, 22, 23, 18, 9, 9, 18, 16,
    #                       # 17, 22, 23, 16, 17, 26, 16, 17, 18, 16,
    #                       # 17, 18, 16, 17, 16, 17, 16, 17, 18, 18, 16,
    #                       # 17, 16, 17, 18, 16, 17, 16, 17, 16, 17, 26
    #                       ])

    pptx_copy.copy_slides()


main()
