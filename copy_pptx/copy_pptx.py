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
    # Открываем исходный PPTX файл как ZIP-архив
    source_folder = f"{script_location}/source_pptx_extracted"
    os.makedirs(source_folder, exist_ok=True)

    with zipfile.ZipFile(source_pptx, 'r') as source_zip:
        # Извлекаем содержимое исходного PPTX файла
        source_zip.extractall(source_folder)

    # Открываем целевой PPTX файл как ZIP-архив
    with zipfile.ZipFile(target_pptx, "a") as target_zip:
        for slide_num in slides_to_copy:
            # Формируем путь к XML-файлу слайда в исходном PPTX файле
            slide_xml_path = f"{source_folder}/ppt/slides/slide{slide_num}.xml"
            # Добавляем XML-файл слайда в целевой PPTX файл
            target_zip.write(slide_xml_path, f"ppt/slides/slide{slide_num}.xml")

            # Копирование медиафайлов, связанных с текущим слайдом
            media_path = f"{source_folder}/ppt/media/"
            for file in os.listdir(media_path):
                if file.startswith(f"slide{slide_num}"):
                    target_zip.write(os.path.join(media_path, file), f"ppt/media/{file}")

            # Копирование замечаний к текущему слайду
            notes_path = f"{source_folder}/ppt/notesSlides/"
            for file in os.listdir(notes_path):
                if file.startswith(f"slide{slide_num}"):
                    target_zip.write(os.path.join(notes_path, file), f"ppt/notesSlides/{file}")

        # Добавляем остальные необходимые файлы из исходной презентации
        for file in source_zip.namelist():
            if file.startswith("ppt/") and file != "ppt/slides/_rels":
                target_zip.write(os.path.join(source_folder, file), file)

    # Удаляем временную папку с извлеченным содержимым исходного PPTX файла
    shutil.rmtree(source_folder)