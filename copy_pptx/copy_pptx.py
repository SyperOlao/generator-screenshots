import io
from pptx import Presentation
from pathlib import Path
from pptx.util import Inches
import warnings
import copy


warnings.filterwarnings("ignore", category=UserWarning)


script_location = Path(__file__).absolute().parent

async def copy_pptx():
    new_presentation = Presentation()
    path_to_template = f"{script_location}/template.pptx"
    source_presentation = Presentation(path_to_template)
    new_presentation = copy_slide_from_external_prs(source_presentation, new_presentation)
    new_presentation.save(f"{script_location}/res.pptx")
    

def copy_slide_from_external_prs(source_presentation, new_presentation):
    source_slide = source_presentation.slides[0]

    # Создать новый слайд в новой презентации, используя макет исходного слайда
    new_slide = new_presentation.slides.add_slide(source_slide.slide_layout)

    # Копировать все элементы с исходного слайда на новый слайд
    for source_shape in source_slide.shapes:
        if source_shape.has_text_frame:
            # Копирование текста
            text = source_shape.text_frame.text
            left = source_shape.left
            top = source_shape.top
            width = source_shape.width
            height = source_shape.height
            new_shape = new_slide.shapes.add_textbox(left, top, width, height)
            new_shape.text_frame.text = text
        elif source_shape.shape_type == 13: # 13 означает изображение
            # Копирование изображения
    
            image = source_shape.image
            image_file = io.BytesIO(image.blob)
            left = source_shape.left
            top = source_shape.top
            width = source_shape.width
            height = source_shape.height
            new_shape = new_slide.shapes.add_picture(image_file, left, top, width=width, height=height)
    
    return new_presentation