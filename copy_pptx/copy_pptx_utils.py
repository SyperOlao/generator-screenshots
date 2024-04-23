import os
import uuid

from config.config import logger
import glob
import re
from io import StringIO
import xml.etree.ElementTree as ET
from lxml import etree
import shutil
import random
import zipfile

from copy_pptx.search_dif_tools import dif_dir
import random


class UniqueHexGenerator:
    def __init__(self):
        self.length = 12
        self.min_val = 300
        self.max_val = 5000
        self.generated = set()

    def generate_unique_hex(self):
        while True:
            random_int = random.randint(0, 16**self.length - 1)
            hex_str = hex(random_int)[2:].upper().zfill(self.length)
            if hex_str not in self.generated:
                self.generated.add(hex_str)
                return str(hex_str)


class CopyPptxUtils:
    @staticmethod
    def save_pptx_as_folder(path_to_pptx, folder_name):
        os.makedirs(folder_name, exist_ok=True)
        with zipfile.ZipFile(path_to_pptx, 'r') as source_zip:
            source_zip.extractall(folder_name)

    @staticmethod
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

    @staticmethod
    def change_notes_slides(path_to_lib, index):
        CopyPptxUtils.change_a_t_notes_slides(path_to_lib, index)
        CopyPptxUtils.change_file_index(path_to_lib, index)
        notes_slides_rels = CopyPptxUtils.change_file_index_rels(path_to_lib, index)
        tree = etree.parse(notes_slides_rels)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(notes_slides_rels)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        for notes_rel in relationship_elements:
            pattern = r'../slides/slide(\d+)\.xml'
            r_num = CopyPptxUtils.extract_slide_numbers(notes_rel.get('Target'), pattern)
            if r_num != index \
                    and r_num is not None:
                notes_rel.set('Target', f'../slides/slide{index}.xml')
        tree.write(notes_slides_rels, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def change_a_t_notes_slides(path, index):
        tree = etree.parse(path)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(path)
        if 'a' not in namespaces:
            return
        a_t = root.find('.//a:t', namespaces=namespaces)
        if a_t is None:
            return
        a_t.text = str(index)

        tree.write(path, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def generate_hex_string():
        group1 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group2 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group3 = ''.join(random.choices('0123456789ABCDEF', k=4))
        group4 = ''.join(random.choices('0123456789ABCDEF', k=12))
        formatted_hex_string = f"{group1}-{group2}-{group3}-{group4}"
        return formatted_hex_string

    @staticmethod
    def generate_hex(n):
        random_int = random.randint(0, 16 ** n - 1)
        hex_str = hex(random_int)[2:].upper()
        return str(hex_str.zfill(n))

    @staticmethod
    def change_chart_id(path_to_chart, hex_generator):
        tree = etree.parse(path_to_chart)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(path_to_chart)
        new_num = dict()
        try:
            unique_ids = root.findall('.//c16:uniqueId', namespaces=namespaces)

            for uid in unique_ids:

                _, _, old_uid = str(uid.get('val')).partition('-')
                if old_uid not in new_num:
                    print("old_uid", uid.get('val'))
                    new_num[old_uid] = hex_generator.generate_unique_hex()

                curr_id = str(uid.get('val')).split('-')[0][1::]

                new_list = str(uid.get('val')).replace('{', '').replace('}', '').split('-')[1:-1]

                result = '-'.join(new_list) + '-' + new_num[old_uid]

                uid.set('val', "{" + f"{curr_id}-{result}" + "}")

            for i in new_num:
                print(new_num[i])
            tree.write(path_to_chart, pretty_print=True, xml_declaration=True, encoding='utf-8')
        except Exception as err:
            logger.warning(err)

    @staticmethod
    def get_last_index(path, pattern):
        files_to_index = glob.glob(os.path.join(path, pattern))
        num = []
        for file in files_to_index:
            num.append(int(CopyPptxUtils.get_number_from_str(file)[0]))

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
            new_name = str(CopyPptxUtils.replace_number(excel_name,
                                                        str(int(embedding_index) - 1)))
        else:
            new_name = excel_name.split('.')[0] + str(int(embedding_index) - 1) + '.xlsx'

        if new_name is None:
            return str(excel_name).join('')
        return new_name

    @staticmethod
    def delete_files_from_folder(slides_path, pattern):
        CopyPptxUtils.delete_all_files(slides_path, pattern)
        CopyPptxUtils.delete_all_files(slides_path + '/_rels', pattern)

    @staticmethod
    def delete_all_files(slides_path, pattern):
        files_to_delete = glob.glob(os.path.join(slides_path, pattern))
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
            except OSError as e:
                logger.warning(f"Error of deleting file '{file_path}': {e}")

    @staticmethod
    def extract_slide_numbers(text, pattern):
        matches = re.findall(pattern, text)
        slide_numbers = [int(match) for match in matches]
        if len(slide_numbers) == 0:
            return None

        return slide_numbers[0]

    @staticmethod
    def extract_before_first_number(s):
        match = re.search(r'^([^\d]*)', s)
        if match:
            return match.group(1)
        else:
            return s

    @staticmethod
    def change_file_index_rels(slides_path, new_index):
        slides_path_2 = slides_path.rsplit('/', 1)[0] + '/_rels'
        file = slides_path_2 + '/' + slides_path.rsplit('/', 1)[1] + '.rels'
        temp_slides_path = slides_path_2 + '/temp'
        CopyPptxUtils.create_a_dir(temp_slides_path)
        tree = etree.parse(file)
        new_file_number = re.sub(r'\d+', str(new_index), str(slides_path.rsplit('/', 1)[1])) + '.rels'
        result_path = f"{temp_slides_path}/{new_file_number}"
        tree.write(result_path, pretty_print=True, xml_declaration=True, encoding='utf-8')
        return result_path

    @staticmethod
    def change_file_index(slides_path, new_index):
        temp_slides_path = slides_path.rsplit('/', 1)[0] + "/temp"
        tree = etree.parse(slides_path)
        CopyPptxUtils.create_a_dir(temp_slides_path)
        new_file_number = re.sub(r'\d+', str(new_index), str(slides_path.rsplit('/', 1)[1]))
        tree.write(temp_slides_path + "/" + new_file_number, pretty_print=True, xml_declaration=True, encoding='utf-8')

    @staticmethod
    def move_file(file_path, new_index):
        temp_slides_path = file_path.rsplit('/', 1)[0] + "/temp"
        CopyPptxUtils.create_a_dir(temp_slides_path)
        new_file_number = re.sub(r'\d+', str(new_index), str(file_path.rsplit('/', 1)[1]))
        new_file_path = f"{temp_slides_path}/{new_file_number}"
        shutil.copy2(file_path, new_file_path)

    @staticmethod
    def rename_and_move_file(old_path: str, new_name: str, new_directory):
        if not os.path.isfile(old_path):
            logger.warning(f"File {old_path} is not found.")
            return
        new_path = new_directory + "/" + new_name
        CopyPptxUtils.create_a_dir(new_directory)

        shutil.copy2(old_path, new_path)
        # logger.info(f"File has been renamed and moved: {old_path} -> {new_path}")

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
        CopyPptxUtils.move_files(source_folder)
        CopyPptxUtils.move_files(source_folder + '/_rels')

    @staticmethod
    def move_files(source_folder):
        temp_folder = source_folder + "/temp"
        if not os.path.exists(temp_folder):
            pass
            # logger.warning(f"Source folder '{source_folder}' can not be find.")
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

    @staticmethod
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
                                    logger.info(f"Word '{word}' found in file: {file_path}, line {line_number}:")
                                    logger.info(line.strip())
                    except Exception as e:
                        logger.warning(f"Error reading file {file_path}: {e}")

    @staticmethod
    def compare_dir(dir1, dir2):
        for root, dirs, files in os.walk(dir1):
            for file in files:
                if file.endswith('.xml') or file.endswith('.rels'):
                    file_path_dir1 = os.path.join(root, file)
                    file_path_dir2 = os.path.join(str(root).replace(dir1, dir2), file)
                    try:
                        with open(file_path_dir2, "r") as file_2:
                            pass
                        tree1 = etree.parse(file_path_dir1)
                        tree1.write(file_path_dir1, pretty_print=True, encoding='utf-8')
                        tree2 = etree.parse(file_path_dir2)
                        tree2.write(file_path_dir2, pretty_print=True, encoding='utf-8')
                        try:
                            dif_dir(file_path_dir1, file_path_dir2)
                        except Exception as e:
                            logger.warning(f"Error reading file: {e}")
                    except FileNotFoundError:
                        print(f"Файл {file_path_dir2} не существует.\n")
