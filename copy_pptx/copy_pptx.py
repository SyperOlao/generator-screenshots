import re
from io import StringIO

from pptx import Presentation
from pathlib import Path
import warnings
import zipfile
import xml.etree.ElementTree as ET
import os
from pprint import pprint
from lxml import etree

warnings.filterwarnings("ignore", category=UserWarning)

script_location = Path(__file__).absolute().parent


def copy_pptx():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)

    path_to_source = f"{script_location}/template.pptx"
    source_presentation = Presentation(path_to_source)

    copy_slides(path_to_source, path_to_new, [1, 2, 4, 7, 27, 21])
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

        for file in source_zip.namelist():
            if file.startswith("ppt/") \
                    and not file.startswith("ppt/notesSlides") \
                    and not file.startswith("ppt/slides") \
                    and not file.startswith("ppt/slideLayouts"):
                target_zip.write(os.path.join(source_folder, file), file)

        copy_all_files(source_zip, target_zip, source_folder)
    # shutil.rmtree(source_folder)


def working_with_xml(source_folder, slides_to_copy):
    root_pptx_xml = f"{source_folder}/ppt/presentation.xml"

    tree = etree.parse(root_pptx_xml)

    root = tree.getroot()
    temp_list = [x + 2 for x in slides_to_copy]
    my_namespaces = dict([
        node for _, node in ET.iterparse(StringIO(str(ET.tostring(root), encoding='utf-8')), events=['start-ns'])])
    pprint(my_namespaces)
    sld_id_lst = root.xpath('//ns0:sldIdLst', namespaces=my_namespaces)
    sld_ids = root.xpath('//ns0:sldId', namespaces=my_namespaces)
    to_remove = []
    for sldId in sld_ids:
        r_id = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        r_num = int(r_id.replace('rId', ''))
        if r_num not in temp_list:
            to_remove.append(sldId)

    for sldId in to_remove:
        parent = sldId.getparent()
        if parent is not None:
            parent.remove(sldId)

    tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')


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
                and not file.startswith("ppt/notesSlides") \
                and not file.startswith("ppt/slides") \
                and not file.startswith("ppt/slideLayouts"):
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
            target_zip.write(os.path.join(source_folder, file), file)


copy_pptx()
