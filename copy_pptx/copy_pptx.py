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

warnings.filterwarnings("ignore", category=UserWarning)

script_location = Path(__file__).absolute().parent


def search_string_in_xml_files_recursive(folder_path, search_string):
    # Используем os.walk для рекурсивного обхода папок
    for root, dirs, files in os.walk(folder_path):
        # Фильтруем список файлов, оставляя только XML файлы
        xml_files = [file for file in files if file.endswith('.xml')]

        # Перебираем XML файлы
        for xml_file in xml_files:
            file_path = os.path.join(root, xml_file)

            # Парсим XML файл
            try:
                tree = etree.parse(file_path)
            except etree.XMLSyntaxError:
                print(f"Ошибка при парсинге файла {file_path}. Пропускаем.")
                continue

            # Преобразуем дерево элементов в строку
            xml_string = etree.tostring(tree, encoding='unicode')

            # Ищем строку в XML строке
            if search_string in xml_string:
                print(f"Строка найдена в файле {file_path}")

        # Пример использования
        folder_path = f"{script_location}/source_pptx_extracted/ppt"
        search_string = 'slide'
        search_string_in_xml_files_recursive(folder_path, search_string)


def copy_pptx():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)

    path_to_source = f"{script_location}/template.pptx"

    copy_slides(path_to_source, path_to_new, [1, 2, 5, 8, 4, 7, 27, 21])


def copy_slides(source_pptx, target_pptx, slides_to_copy):
    source_folder = f"{script_location}/source_pptx_extracted"
    os.makedirs(source_folder, exist_ok=True)

    with zipfile.ZipFile(source_pptx, 'r') as source_zip:
        source_zip.extractall(source_folder)

    with (zipfile.ZipFile(target_pptx, "w") as target_zip):
        working_with_xml(source_folder, slides_to_copy)

        copy_all_files(source_zip, target_zip, source_folder)

    shutil.rmtree(source_folder)


def working_with_xml(source_folder, slides_to_copy):
    root_pptx_xml = f"{source_folder}/ppt/presentation.xml"
    root_pptx_xml_rels = f"{source_folder}/ppt/_rels/presentation.xml.rels"
    root_content_types = f"{source_folder}/[Content_Types].xml"
    doc_props = f"{source_folder}/docProps/app.xml"
    change_doc_props(doc_props, slides_to_copy)
    change_root_pptx_xml(root_pptx_xml, slides_to_copy)
    change_root_pptx_xml_rels(root_pptx_xml_rels, slides_to_copy)
    change_root_context_type(root_content_types, slides_to_copy)
    i = 0
    slides_path = f"{source_folder}/ppt/slides"
    temp_slides_path = f"{source_folder}/ppt/slides/temp"

    slides_path_rels = f"{source_folder}/ppt/slides/_rels"
    temp_slides_path_rels = f"{source_folder}/ppt/slides/_rels/temp"

    slides_path_note = f"{source_folder}/ppt/notesSlides"
    temp_slides_path_note = f"{source_folder}/ppt/notesSlides/temp"

    slides_path_note_rels = f"{source_folder}/ppt/notesSlides/_rels"
    temp_slides_path_note_rels = f"{source_folder}/ppt/notesSlides/_rels/temp"

    for slide in slides_to_copy:
        i += 1
        change_slide_pptx(f"{slides_path}/slide{slide}.xml", f"{temp_slides_path}/slide{i}.xml", temp_slides_path)
        change_rels_file(f"{slides_path_rels}/slide{slide}.xml.rels", f"{temp_slides_path_rels}/slide{i}.xml.rels",
                         temp_slides_path_rels, slide, i)
        change_slide_pptx(f"{slides_path_note}/notesSlide{slide}.xml", f"{temp_slides_path_note}/notesSlide{i}.xml",
                          temp_slides_path_note)

        change_rels_file(f"{slides_path_note_rels}/notesSlide{slide}.xml.rels",
                         f"{temp_slides_path_note_rels}/notesSlide{i}.xml.rels",
                         temp_slides_path_note_rels, slide, i)

    delete_all_slides(slides_path, 'slide*')
    delete_all_slides(slides_path_rels, 'slide*')
    delete_all_slides(slides_path_note, 'notesSlide*')
    delete_all_slides(slides_path_note_rels, 'notesSlide*')

    move_slides(temp_slides_path, slides_path)
    move_slides(temp_slides_path_rels, slides_path_rels)
    move_slides(temp_slides_path_note, slides_path_note)
    move_slides(temp_slides_path_note_rels, slides_path_note_rels)


def delete_all_slides(slides_path, pattern):
    files_to_delete = glob.glob(os.path.join(slides_path, pattern))
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
        except OSError as e:
            logging.warning(f"Error of deleting file '{file_path}': {e}")


def move_slides(source_folder, destination_folder):
    if not os.path.exists(source_folder):
        logging.warning(f"Source folder '{source_folder}' can not be find.")
    else:
        for filename in os.listdir(source_folder):
            file_path = os.path.join(source_folder, filename)
            if os.path.isfile(file_path):
                shutil.move(file_path, destination_folder)

        shutil.rmtree(source_folder)


def change_doc_props(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    namespaces = get_name_spaces_by_filepath(root_pptx_xml)

    num = root.find('.//Slides', namespaces=namespaces)
    print(num.text)
    current_slides_count = len(slides_to_copy)
    num.text = str(current_slides_count)
    relationship_elements = root.findall('.//vt:lpstr', namespaces=namespaces)
    i = 0
    for rel in relationship_elements:
        if "PowerPoint" in rel.text:
            if i > current_slides_count:
                delete_child(rel)
            i += 1
    tree.write(root_pptx_xml)


def change_slide_pptx(slides_path, slide_xml_path_new, temp_slides_path):
    tree = etree.parse(slides_path)
    if not os.path.exists(temp_slides_path):
        os.makedirs(temp_slides_path)
    tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')


def change_rels_file(slides_path, slide_xml_path_new, temp_slides_path, old_num, new_num):
    change_slide_pptx(slides_path, slide_xml_path_new, temp_slides_path)
    tree = etree.parse(slide_xml_path_new)
    root = tree.getroot()

    namespaces = get_name_spaces_by_filepath(slide_xml_path_new)

    relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
    for rel in relationship_elements:
        pattern = r'../slides/slide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('Target'), pattern)

        if r_num != new_num \
                and r_num is not None:
            rel.set('Target', f'../slides/slide{new_num}.xml')

    for rel in relationship_elements:
        pattern = r'../notesSlides/notesSlide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('Target'), pattern)
        if r_num != new_num \
                and r_num is not None:
            rel.set('Target', f'../notesSlides/notesSlide{new_num}.xml')

    tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')


def extract_slide_numbers(text, pattern):
    matches = re.findall(pattern, text)
    slide_numbers = [int(match) for match in matches]
    if len(slide_numbers) == 0:
        return None

    return slide_numbers[0]


def change_root_pptx_xml_rels(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    namespaces = get_name_spaces_by_filepath(root_pptx_xml)

    relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
    for rel in relationship_elements:
        pattern = r'slides/slide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('Target'), pattern)
        new_slides = [x + 1 for x in range(len(slides_to_copy))]

        if r_num not in new_slides \
                and r_num is not None:
            delete_child(rel)

    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


def delete_child(rel):
    parent = rel.getparent()
    if parent is not None:
        parent.remove(rel)


def change_root_context_type(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    namespaces = get_name_spaces_by_filepath(root_pptx_xml)

    new_slides = [x + 1 for x in range(len(slides_to_copy))]

    relationship_elements = root.findall('.//Override', namespaces=namespaces)
    for rel in relationship_elements:
        pattern = r'/ppt/slides/slide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('PartName'), pattern)

        if r_num not in new_slides \
                and r_num is not None:
            delete_child(rel)
        r_num = extract_slide_numbers(rel.get('PartName'), r'/ppt/notesSlides/notesSlide(\d+)\.xml')
        if r_num not in new_slides \
                and r_num is not None:
            delete_child(rel)
    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


def change_root_pptx_xml(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    namespaces = get_name_spaces(root)

    sld_ids = root.xpath('//ns0:sldId', namespaces=namespaces)

    for sldId in sld_ids[len(slides_to_copy)::]:
        parent = sldId.getparent()
        if parent is not None:
            parent.remove(sldId)

    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


def get_name_spaces(root):
    return dict([
        node for _, node in ET.iterparse(StringIO(str(ET.tostring(root), encoding='utf-8')), events=['start-ns'])])


def get_name_spaces_by_filepath(filepath):
    return dict([node for _, node in ET.iterparse(filepath,
                                                  events=['start-ns'])])

def copy_all_files(source_zip, target_zip, source_folder):
    """
    Adding common files for pptx from source_zip pptx
    :param source_zip: source pptx opened like zip
    :param target_zip: target pptx opened like zip
    :param source_folder: - place where pptx temp extracted folder is located
    :return:
    """
    for file in source_zip.namelist():
        try:
            target_zip.write(os.path.join(source_folder, file), file)
        except OSError as e:
            logging.warning(e)


copy_pptx()
