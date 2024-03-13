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


def copy_pptx():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)

    path_to_source = f"{script_location}/template.pptx"
    source_presentation = Presentation(path_to_source)

    copy_slides(path_to_source, path_to_new, [1, 2, 5, 8, 4, 7, 27, 21])
    # new_presentation.save(path_to_new)


def copy_slides(source_pptx, target_pptx, slides_to_copy):
    source_folder = f"{script_location}/source_pptx_extracted"
    os.makedirs(source_folder, exist_ok=True)

    with zipfile.ZipFile(source_pptx, 'r') as source_zip:
        source_zip.extractall(source_folder)

    with (zipfile.ZipFile(target_pptx, "a") as target_zip):

        working_with_xml(source_folder, slides_to_copy)

        for slide_num in slides_to_copy:
            slide_xml_path = f"{source_folder}/ppt/slides"
            target_zip.write(slide_xml_path, f"{slide_xml_path}/slide{slide_num}.xml")

            slide_layout_xml_path = f"{source_folder}/ppt/slideLayouts"
            target_zip.write(slide_layout_xml_path, f"{slide_layout_xml_path}/slideLayout{slide_num}.xml")

            media_path = f"{source_folder}/ppt/media/"
            for file in os.listdir(media_path):
                target_zip.write(os.path.join(media_path, file), f"ppt/media/{file}")

            notes_path = f"{source_folder}/ppt/notesSlides/"
            for file in os.listdir(notes_path):
                if file == f"notesSlide{slide_num}.xml":
                    target_zip.write(os.path.join(notes_path, file), f"ppt/notesSlides/{file}")

        # copy_solely_necessary_files(source_zip, target_zip, source_folder)

        copy_all_files(source_zip, target_zip, source_folder)
    # shutil.rmtree(source_folder)


def working_with_xml(source_folder, slides_to_copy):
    root_pptx_xml = f"{source_folder}/ppt/presentation.xml"
    root_pptx_xml_rels = f"{source_folder}/ppt/_rels/presentation.xml.rels"
    root_content_types = f"{source_folder}/[Content_Types].xml"
    change_root_pptx_xml(root_pptx_xml, slides_to_copy)
    change_root_pptx_xml_rels(root_pptx_xml_rels, slides_to_copy)
    change_root_context_type(root_content_types, slides_to_copy)
    i = 0
    slides_path = f"{source_folder}/ppt/slides"
    temp_slides_path = f"{source_folder}/ppt/slides/temp"

    slides_path_rels = f"{source_folder}/ppt/slides/_rels"
    temp_slides_path_rels = f"{source_folder}/ppt/slides/_rels/temp"
    for slide in slides_to_copy:
        i += 1
        change_slide_pptx(slides_path, temp_slides_path, slide, i)
        change_rels_file(slides_path_rels, temp_slides_path_rels, slide, i)
    delete_all_slides(slides_path)
    delete_all_slides(slides_path_rels)

    move_slides(temp_slides_path, slides_path)
    move_slides(temp_slides_path_rels, slides_path_rels)


def delete_all_slides(slides_path):
    files_to_delete = glob.glob(os.path.join(slides_path, 'slide*'))
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


def change_slide_pptx(slides_path, temp_slides_path, slide_num_old, slide_num_new):
    slide_xml_path = f"{slides_path}/slide{slide_num_old}.xml"
    tree = etree.parse(slide_xml_path)
    root = tree.getroot()
    namespaces = get_name_spaces_by_filepath(slide_xml_path)
    slide_id = root.find('.//a:t', namespaces=namespaces)
    if slide_id is not None:
        slide_id.text = str(slide_num_new)

    # blip_element = root.find('.//a:blip', namespaces=namespaces)
    # r = "{" + namespaces['r'] + "}"
    #
    # if blip_element is not None:
    #     new_embed_value = f'rId{slide_num_new + 2}'
    #
    #     blip_element.set(f"{r}embed", new_embed_value)

    if not os.path.exists(temp_slides_path):
        os.makedirs(temp_slides_path)
    slide_xml_path_new = f"{temp_slides_path}/slide{slide_num_new}.xml"
    tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')


def change_rels_file(slides_path, temp_slides_path, slide_num_old, slide_num_new):
    slide_xml_path = f"{slides_path}/slide{slide_num_old}.xml.rels"
    tree = etree.parse(slide_xml_path)
    if not os.path.exists(temp_slides_path):
        os.makedirs(temp_slides_path)
    slide_xml_path_new = f"{temp_slides_path}/slide{slide_num_new}.xml.rels"
    tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')


def extract_slide_numbers(text, pattern):
    # Регулярное выражение для поиска строк, начинающихся на 'slides/slide', за которыми следуют цифры

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
        pattern = r'/ppt/slides/slide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('Target'), pattern)

        if r_num not in slides_to_copy \
                and r_num is not None:

            parent = rel.getparent()
            if parent is not None:
                parent.remove(rel)

    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


def change_root_context_type(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    namespaces = get_name_spaces_by_filepath(root_pptx_xml)

    new_slides = [x+1 for x in range(len(slides_to_copy))]

    pprint(new_slides)
    relationship_elements = root.findall('.//Override', namespaces=namespaces)
    for rel in relationship_elements:
        pattern = r'slides/slide(\d+)\.xml'
        r_num = extract_slide_numbers(rel.get('PartName'), pattern)

        if r_num not in new_slides \
                and r_num is not None:

            parent = rel.getparent()
            if parent is not None:
                parent.remove(rel)



    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


def change_root_pptx_xml(root_pptx_xml, slides_to_copy):
    tree = etree.parse(root_pptx_xml)
    root = tree.getroot()

    # TODO:: Add step by master slides
    slides_to_copy_with_step = [x + 2 for x in slides_to_copy]
    namespaces = get_name_spaces(root)

    # sld_id_lst = root.xpath('//ns0:sldIdLst', namespaces=namespaces)[0]
    sld_ids = root.xpath('//ns0:sldId', namespaces=namespaces)

    #
    # for sldId in sld_ids:
    #     r_id = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    #     r_num = int(r_id.replace('rId', ''))
    #     if r_num in slides_to_copy_with_step:
    #         res_pptx[r_num] = sldId

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


def copy_solely_necessary_files(source_zip, target_zip, source_folder):
    """
    Adding addition files for pptx from source_zip pptx except some unnecessary files
    :param source_zip: source pptx opened like zip
    :param target_zip: target pptx opened like zip
    :param source_folder: place where pptx temp extracted folder is located
    :return:
    """
    for file in source_zip.namelist():
        if file.startswith("ppt/") \
                and not file.startswith("ppt/slides") \
                and not file.startswith("ppt/_rels"):
            target_zip.write(os.path.join(source_folder, file), file)


def copy_all_files(source_zip, target_zip, source_folder):
    """
    Adding common files for pptx from source_zip pptx
    :param source_zip: source pptx opened like zip
    :param target_zip: target pptx opened like zip
    :param source_folder: - place where pptx temp extracted folder is located
    :return:
    """
    for file in source_zip.namelist():
        if file.startswith("ppt/"):
            try:
                target_zip.write(os.path.join(source_folder, file), file)
            except OSError as e:
                logging.warning(e)


copy_pptx()
