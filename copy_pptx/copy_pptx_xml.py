from config.config import logger
from pptx import Presentation
from pathlib import Path
import zipfile
import os
from lxml import etree
import shutil
from copy_pptx_utils import CopyPptxUtils

script_location = Path(__file__).absolute().parent


class CopyPptx:
    _source_folder = f"{script_location}/{CopyPptxUtils.generate_random_string()}"
    _target_indexes = dict()
    _repeated_indexes = dict()
    _styles = []
    _font_ids = dict()
    _num_of_words = 0
    _num_of_paragraphs = 0

    def __init__(self, path_to_source_pptx, path_to_new_pptx, slides_to_copy):
        self.path_to_source = path_to_source_pptx
        self.path_to_new = path_to_new_pptx
        self.slides_to_copy = slides_to_copy
        self._len_master_id = 0
        self._get_repeated_indexes(slides_to_copy)

    def copy_slides(self):
        """
        Major method to coping slides from source pptx to
        new pptx

        Work principal is generating and changing root xml
        and copying them to new pptx opened as folder
        :return: None
        """
        os.makedirs(self._source_folder, exist_ok=True)

        with zipfile.ZipFile(self.path_to_source, 'r') as source_zip:
            source_zip.extractall(self._source_folder)

        with (zipfile.ZipFile(self.path_to_new, "w") as target_zip):
            self._working_with_xml()

            self._copy_all_files(target_zip)

        shutil.rmtree(self._source_folder)

    def _working_with_xml(self):
        """
        Method of changing and creating xml and copying them to
        the new presentation folder
        :return: None
        """

        slides_path = f"{self._source_folder}/ppt/slides"
        self._change_root_pptx_xml()
        self._change_root_pptx_xml_rels()
        i = 0

        for slide in self.slides_to_copy:
            i += 1
            self._update_doc_props(f"{slides_path}/slide{slide}.xml")
            CopyPptxUtils.change_file_index(f"{slides_path}/slide{slide}.xml", i)
            self._change_rels_file(f"{slides_path}/slide{slide}.xml", i, slide)

        self._change_doc_props()
        self._change_root_context_type()
        self._delete_and_move_files()

    def _update_doc_props(self, slides_path):
        tree = etree.parse(slides_path)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(slides_path)
        a_t = root.findall('.//a:t', namespaces=namespaces)

        for e in a_t:
            self._num_of_words += len(e.text.split())
        self._num_of_words += 1
        self._num_of_paragraphs += len(root.findall('.//a:p', namespaces=namespaces))
        self._num_of_paragraphs += 1

    def _change_doc_props(self):
        root_pptx_xml = f"{self._source_folder}/docProps/app.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(root_pptx_xml)

        slides = root.find('.//Slides', namespaces=namespaces)
        notes = root.find('.//Notes', namespaces=namespaces)
        paragraphs = root.find('.//Paragraphs', namespaces=namespaces)
        words = root.find('.//Words', namespaces=namespaces)
        current_slides_count = len(self.slides_to_copy)
        paragraphs.text = str(self._num_of_paragraphs)
        words.text = str(self._num_of_words)
        slides.text = str(current_slides_count)
        notes.text = str(current_slides_count)
        relationship_elements = root.findall('.//vt:lpstr', namespaces=namespaces)
        i4 = root.findall('.//vt:i4', namespaces=namespaces)
        if i4 is not None:
            i4[-1].text = str(current_slides_count)
        parent = None
        text = ""

        for rel in relationship_elements:
            if "PowerPoint" in rel.text:
                text = rel.text
                parent = rel.getparent()
                CopyPptxUtils.delete_child(rel)

        amount_slices = len(self.slides_to_copy)
        for i in range(amount_slices):
            a = etree.SubElement(parent, "{" + namespaces['vt'] + "}lpstr")
            a.text = str(text)

        for vector in root.findall('.//vt:vector', namespaces=namespaces):
            base_type = vector.get('baseType')
            if base_type == 'lpstr':
                z = vector.getchildren()
                vector.set('size', str(len(z)))

        tree.write(root_pptx_xml)

    def _change_root_context_type(self):
        """
        [Content_Types].xml is the file with information of all
        external files
        :return: None
        """
        root_pptx_xml = f"{self._source_folder}/[Content_Types].xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(root_pptx_xml)
        content_type = dict()
        relationship_elements = root.findall('.//Override', namespaces=namespaces)
        relations = None
        for rel in relationship_elements:
            target_type = str(str(rel.get('ContentType')).split('.')[-1].split('+')[0]).lower()

            if target_type == 'slidelayout':
                continue

            content_type[target_type] = {'ct': str(rel.get('ContentType')), 'pt': str(rel.get('PartName'))}

            if target_type in ['slide', 'chart', 'notesslide', 'chartstyle', 'chartcolorstyle']:
                relations = rel.getparent()
                CopyPptxUtils.delete_child(rel)

        for i in range(len(self.slides_to_copy)):
            etree.SubElement(relations, "Override",
                             {
                                 "PartName": CopyPptxUtils.replace_number(f"{content_type['slide']['pt']}",
                                                                          str(i + 1)),
                                 "ContentType": f"{content_type['slide']['ct']}"
                             })

        for target_type in self._target_indexes:
            if target_type.lower() not in content_type:
                continue

            for i in range(self._target_indexes[target_type]):
                etree.SubElement(relations, "Override",
                                 {
                                     "PartName": CopyPptxUtils.replace_number(
                                         f"{content_type[target_type.lower()]['pt']}",
                                         str(i + 1)),
                                     "ContentType": f"{content_type[target_type.lower()]['ct']}"
                                 })

        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_root_pptx_xml_rels(self):
        """
        Change relations links to new
        .rels xml file is files with all relations.
        :return:
        """
        root_pptx_xml = f"{self._source_folder}/ppt/_rels/presentation.xml.rels"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(root_pptx_xml)
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
            index = self._len_master_id + i + 1
            etree.SubElement(relations, "Relationship",
                             {'Id': f'rId{index}',
                              "Type": f'{all_type}',
                              "Target": f'slides/slide{i + 1}.xml'})

        max_value = max(self._font_ids.values(), key=lambda x: int(x[3:]))
        index = int(str(CopyPptxUtils.get_number_from_str(max_value)[0]))
        for rel in relationship_elements:
            target_type = str(rel.get('Type')).split('/')[-1]
            if target_type == 'slide':
                continue

            if target_type == 'font' \
                    or target_type == 'notesMaster' \
                    or target_type == 'slideMaster':
                id = rel.get("Id")
                if id in self._font_ids:
                    rel.set("Id", f'{self._font_ids[id]}')
            else:
                index += 1
                rel.set("Id", f'rId{index}')
        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_root_pptx_xml(self):
        """
        The presentation.xml is the origin file of pptx contains all
        information of slides and font and their id.
        This method changes ids of new slides and removes or adds new
        slides and links to them new id
        :return:
        """
        root_pptx_xml = f"{self._source_folder}/ppt/presentation.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces(root)
        sld_ids = root.xpath('//ns0:sldId', namespaces=namespaces)
        self._len_master_id = len(root.xpath('//ns0:sldMasterId', namespaces=namespaces))

        for sldId in sld_ids:
            CopyPptxUtils.delete_child(sldId)
        sldIdLst = root.find('ns0:sldIdLst', namespaces=namespaces)
        ids = CopyPptxUtils.generate_ids(len(self.slides_to_copy))
        index = 0
        for i in range(len(self.slides_to_copy)):
            index = i + 1 + self._len_master_id
            etree.SubElement(sldIdLst, "{" + namespaces['ns0'] + "}sldId",
                             {'id': f'{str(ids[i])}',
                              "{" + namespaces['ns1'] + "}id": f'rId{index}'})

        name = "{" + namespaces['ns1'] + "}"
        notes_master_id = root.findall('.//ns0:notesMasterId', namespaces=namespaces)
        for elem in notes_master_id:
            elem_id = elem.get(f'{name}id')
            if elem_id:
                index += 1
                self._font_ids[elem_id] = f'rId{index}'
                elem.set(f"{name}id", f'rId{index}')

        sld_master_id = root.findall('.//ns0:sldMasterId', namespaces=namespaces)
        for elem in sld_master_id:
            elem_id = elem.get(f'{name}id')
            if elem_id:
                self._font_ids[elem_id] = elem_id

        embedded_fonts = root.findall(f'.//ns0:embeddedFont', namespaces=namespaces)
        for embedded_font in embedded_fonts:
            for c in embedded_font.getchildren():
                elem_id = c.get(f'{name}id')
                if elem_id:
                    index += 1
                    self._font_ids[elem_id] = f'rId{index}'
                    c.set(f"{name}id", f'rId{index}')

        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_rels_file(self, slides_path, new_index, old_index):
        slide_xml_path_new = CopyPptxUtils.change_file_index_rels(slides_path, new_index)
        self._deep_change_target_links_rels(slide_xml_path_new, old_index)

    def _deep_change_target_links_rels(self, slide_xml_path_new, old_index):
        tree = etree.parse(slide_xml_path_new)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(slide_xml_path_new)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)
        notes_slides_path = f"{self._source_folder}/ppt/notesSlides/notesSlide"
        if old_index in self._repeated_indexes:
            self._repeated_indexes[old_index] += 1

        for rel in relationship_elements:
            target = str(rel.get('Target'))
            target_type = str(rel.get('Type')).split('/')[-1]
            path_to_rel = target.replace('..', self._source_folder + '/ppt')
            index = self._add_target_indexes(target_type)
            if target_type == 'chart':
                self._change_chart_rels(path_to_rel, index)
                rel.set('Target', f'../charts/chart{index}.xml')

            if target_type == 'notesSlide':
                pattern = r'../notesSlides/notesSlide(\d+)\.xml'
                r_num = CopyPptxUtils.extract_slide_numbers(rel.get('Target'), pattern)
                if r_num != index \
                        and r_num is not None:
                    rel.set('Target', f'../notesSlides/notesSlide{index}.xml')
                CopyPptxUtils.change_slide_id(notes_slides_path, old_index)
                CopyPptxUtils.change_notes_slides(path_to_rel, index)

        tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_chart_rels(self, path_to_chart, index):
        CopyPptxUtils.change_file_index(path_to_chart, index)
        chart_path_rels = CopyPptxUtils.change_file_index_rels(path_to_chart, index)

        tree = etree.parse(chart_path_rels)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(chart_path_rels)
        relationship_elements = root.findall('.//Relationship', namespaces=namespaces)

        for rel in relationship_elements:
            chart_target = str(rel.get('Target'))
            chart_target_type = str(rel.get('Type')).split('/')[-1]
            if chart_target_type == 'package':
                self._change_package(rel, chart_target_type, chart_target)
            if chart_target_type == 'chartStyle':
                self._change_chart_style(rel, chart_target_type, chart_target)
            if chart_target_type == 'chartColorStyle':
                self._change_chart_style(rel, chart_target_type, chart_target)

        tree.write(chart_path_rels, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_chart_style(self, rel, chart_target_type, chart_target):

        chart_path_to_embedding = self._source_folder + '/ppt/charts'
        embedding_index = str(self._add_target_indexes(chart_target_type))

        CopyPptxUtils.move_file(chart_path_to_embedding + '/' + chart_target, embedding_index)
        rel.set('Target', CopyPptxUtils.replace_number(chart_target, embedding_index))

    def _change_package(self, rel, chart_target_type, chart_target):
        embedding_index = str(self._add_target_indexes(chart_target_type))

        chart_path_to_embedding = chart_target.replace('..', self._source_folder + '/ppt')

        new_chart_path = os.path.dirname(chart_path_to_embedding) + '/temp'

        new_name = CopyPptxUtils.get_embedding_name(chart_path_to_embedding, embedding_index)

        CopyPptxUtils.rename_and_move_file(chart_path_to_embedding,
                                           new_name, new_chart_path)

        rel.set('Target', f'../embeddings/{new_name}')

    def _copy_all_files(self, target_zip):
        """
        Adding common files for pptx from source_folder
        :param target_zip: target pptx opened like zip
        :return: none
        """
        for root, dirs, files in os.walk(self._source_folder):
            for file in files:
                source_path = os.path.join(root, file)
                relative_path = os.path.relpath(source_path, self._source_folder)
                try:
                    target_zip.write(source_path, relative_path)
                except OSError as e:
                    logger.warning(e)

    def _add_target_indexes(self, target_type):
        if target_type in self._target_indexes:
            self._target_indexes[target_type] = self._target_indexes.get(target_type) + 1
        else:
            self._target_indexes[target_type] = 1

        return self._target_indexes[target_type]

    def _get_repeated_indexes(self, numbers):
        counts = {}
        for num in numbers:
            if num in counts:
                counts[num] += 1
            else:
                counts[num] = 1
        rep = [num for num, count in counts.items() if count > 1]
        self._repeated_indexes = dict()
        for i in rep:
            self._repeated_indexes[i] = 0

    def _delete_and_move_files(self):
        """
        Deletes useless files and moves generating xml from temp folder
        to source
        :return:
        """
        slides_path = f"{self._source_folder}/ppt/slides"
        slides_path_note = f"{self._source_folder}/ppt/notesSlides"
        slides_path_charts = f"{self._source_folder}/ppt/charts"
        slides_path_embeddings = f"{self._source_folder}/ppt/embeddings"

        CopyPptxUtils.delete_files_from_folder(slides_path, 'slide*')
        CopyPptxUtils.delete_files_from_folder(slides_path_note, 'notesSlide*')
        CopyPptxUtils.delete_files_from_folder(slides_path_charts, 'chart*')
        CopyPptxUtils.delete_files_from_folder(slides_path_charts, 'colors*')
        CopyPptxUtils.delete_files_from_folder(slides_path_charts, 'style*')
        CopyPptxUtils.delete_files_from_folder(slides_path_embeddings, 'Microsoft_Excel_Worksheet*')

        CopyPptxUtils.move_all_files(slides_path)
        CopyPptxUtils.move_all_files(slides_path_note)
        CopyPptxUtils.move_all_files(slides_path_charts)
        CopyPptxUtils.move_files(slides_path_embeddings)


def main():
    new_presentation = Presentation()
    path_to_new = f"{script_location}/res.pptx"
    new_presentation.save(path_to_new)
    path_to_source = f"{script_location}/template.pptx"

    slides_to_copy = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,
                      28, 29, 30, 31, 32, 33, 34, 35, 18, 21, 16, 17, 19, 20, 22, 23, 25, 24, 22, 23, 25, 24, 22, 23,
                      25, 24, 16, 17, 19, 20, 16, 17, 19, 20, 32, 33, 26, 28, 12, 13, 16, 17, 19, 20, 9, 10, 18, 21, 16,
                      17, 19, 20, 9, 10, 18, 21, 16, 17, 19, 20, 16, 17, 19, 20, 16, 17, 19, 20, 16, 17, 19, 20, 22, 23,
                      25, 24, 22, 23, 25, 24, 18, 21, 16, 17, 19, 20, 18, 21, 16, 17, 19, 20, 18, 21, 18, 21, 22, 23,
                      25, 24, 16, 17, 19, 20, 16, 17, 19, 20, 16, 17, 19, 20, 16, 17, 19, 20, 18, 21, 16, 17, 19, 20]

    pptx_copy = CopyPptx(path_to_source, path_to_new,
                         slides_to_copy)

    pptx_copy.copy_slides()


main()
