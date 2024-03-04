from pptx import Presentation
from pathlib import Path
import warnings
import zipfile
import shutil
import os

warnings.filterwarnings("ignore", category=UserWarning)

script_location = Path(__file__).absolute().parent


async def copy_pptx():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)
    path_to_source = f"{script_location}/template.pptx"
    source_presentation = Presentation(path_to_source)

    copy_slides(path_to_source, path_to_new, [1, 2])
    # new_presentation.save(path_to_new)


def copy_slides(source_pptx, target_pptx, slides_to_copy):
    # Открываем исходный PPTX файл как ZIP-архив
    with zipfile.ZipFile(source_pptx, 'r') as source_zip:
        # Извлекаем содержимое исходного PPTX файла
        source_zip.extractall('source_pptx_extracted')

    # Открываем целевой PPTX файл как ZIP-архив
    with zipfile.ZipFile(target_pptx, "a") as target_zip:
        for slide_num in slides_to_copy:
            # Формируем путь к XML-файлу слайда в исходном PPTX файле
            slide_xml_path = f"source_pptx_extracted/ppt/slides/slide{slide_num}.xml"
            # Добавляем XML-файл слайда в целевой PPTX файл
            target_zip.write(slide_xml_path, f"ppt/slides/slide{slide_num}.xml")

            # Копирование медиафайлов, связанных с слайдом
            media_path = f"source_pptx_extracted/ppt/media/"
            for file in os.listdir(media_path):
                source_file = os.path.join(media_path, file)
                target_file = f"ppt/media/{file}"
                if os.path.exists(source_file):
                    target_zip.write(source_file, target_file)

            # Копирование замечаний к слайду
            notes_path = f"source_pptx_extracted/ppt/notesSlides/"
            for file in os.listdir(notes_path):
                source_file = os.path.join(notes_path, file)
                target_file = f"ppt/notesSlides/{file}"
                if os.path.exists(source_file):
                    target_zip.write(source_file, target_file)

    # Удаляем временную папку с извлеченным содержимым исходного PPTX файла
    shutil.rmtree('source_pptx_extracted')
