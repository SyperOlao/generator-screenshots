from pptx import Presentation
from pathlib import Path
import warnings
import zipfile
import xml.etree.ElementTree as ET
import shutil
import os

warnings.filterwarnings("ignore", category=UserWarning)

script_location = Path(__file__).absolute().parent


def copy_pptx():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)

    path_to_source = f"{script_location}/template.pptx"
    source_presentation = Presentation(path_to_source)

    copy_slides(path_to_source, path_to_new, [1, 2])
    # new_presentation.save(path_to_new)


def copy_slides(source_pptx, target_pptx, slides_to_copy):
    source_folder = f"{script_location}/source_pptx_extracted"
    os.makedirs(source_folder, exist_ok=True)

    with zipfile.ZipFile(source_pptx, 'r') as source_zip:
        source_zip.extractall(source_folder)

    with (zipfile.ZipFile(target_pptx, "a") as target_zip):
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

        working_with_xml(source_folder, slides_to_copy)

        for file in source_zip.namelist():
            if file.startswith("ppt/") \
                    and not file.startswith("ppt/notesSlides") \
                    and not file.startswith("ppt/slides") \
                    and not file.startswith("ppt/slideLayouts"):
                target_zip.write(os.path.join(source_folder, file), file)

        # # Добавляем общие файлы и структуры, необходимые для работы презентации
        # for file in source_zip.namelist():
        #     if file.startswith("ppt/"):
        #         target_zip.write(os.path.join(source_folder, file), file)

    # shutil.rmtree(source_folder)

    # for file in source_zip.namelist():
    #     if file.startswith("ppt/") \
    #             and not file.startswith("ppt/notesSlides") \
    #             and not file.startswith("ppt/slides") \
    #             and not file.startswith("ppt/slideLayouts"):
    #         target_zip.write(os.path.join(source_folder, file), file)


def working_with_xml(source_folder, slides_to_copy):
    root_pptx_xml = f"{source_folder}/ppt/presentation.xml"

    tree = ET.parse(root_pptx_xml)

    root = tree.getroot()
    # for child in root:
    #     print(f"Tag: {child.tag}, Attr {child.attrib}")

    elems = root.findall('.//p:sldId',
                         namespaces={'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                                     'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
    # Нахождение элемента для удаления по атрибуту r:id
    to_remove = []
    for sldId in elems:
        r_id = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        r_num = int(r_id.replace('rId', ''))
        if r_num not in slides_to_copy:
            to_remove.append(sldId)

    for sldId in to_remove:
        print(sldId)
        elems.remove(sldId)

    # Сохранение обновлённого XML-документа
    tree.write(root_pptx_xml)


copy_pptx()
