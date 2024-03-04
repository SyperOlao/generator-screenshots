from pptx import Presentation
from pathlib import Path
import warnings
import zipfile
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

    with zipfile.ZipFile(target_pptx, "a") as target_zip:
        for slide_num in slides_to_copy:
            slide_xml_path = f"{source_folder}/ppt/slides/slide{slide_num}.xml"
            target_zip.write(slide_xml_path, f"ppt/slides/slide{slide_num}.xml")

            media_path = f"{source_folder}/ppt/media/"
            for file in os.listdir(media_path):
                if file.startswith(f"slide{slide_num}"):
                    target_zip.write(os.path.join(media_path, file), f"ppt/media/{file}")

            notes_path = f"{source_folder}/ppt/notesSlides/"
            for file in os.listdir(notes_path):
                if file.startswith(f"slide{slide_num}"):
                    target_zip.write(os.path.join(notes_path, file), f"ppt/notesSlides/{file}")

            # Добавляем файлы отношений для слайдов, изображений, видео и т.д.
            for file in source_zip.namelist():
                if file.startswith(f"ppt/slides/_rels/slide{slide_num}.rels") or \
                   file.startswith(f"ppt/media/") or \
                   file.startswith(f"ppt/notesSlides/") or \
                   file.startswith(f"ppt/embeddings/") or \
                   file.startswith(f"ppt/embeddings/_rels/"):
                    target_zip.write(os.path.join(source_folder, file), file)

        # Добавляем общие файлы и структуры, необходимые для работы презентации
        for file in source_zip.namelist():
            if file.startswith("ppt/") and file != "ppt/slides/_rels":
                target_zip.write(os.path.join(source_folder, file), file)

    shutil.rmtree(source_folder)
