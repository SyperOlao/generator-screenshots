import random
from pprint import pprint

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
    source_folder = f"{script_location}/source_pptx_extracted"
    target_indexes = dict()
    repeated_indexes = dict()
    styles = []
    font_ids = dict()
    num_of_words = 0
    num_of_paragraphs = 0

    def __init__(self, path_to_source, path_to_new, slides_to_copy):
        self.path_to_source = path_to_source
        self.path_to_new = path_to_new
        self.slides_to_copy = slides_to_copy
        self.len_master_id = 0
        self.get_repeated_indexes(slides_to_copy)
        for i in range(len(slides_to_copy)):
            print("i: ", i + 1, " ", slides_to_copy[i])

    def copy_slides(self):
        shutil.rmtree(self.source_folder)
        os.makedirs(self.source_folder, exist_ok=True)

        with zipfile.ZipFile(self.path_to_source, 'r') as source_zip:
            source_zip.extractall(self.source_folder)

        with (zipfile.ZipFile(self.path_to_new, "w") as target_zip):
            self._working_with_xml()

            self._copy_all_files(target_zip)

        # shutil.rmtree(self.source_folder)

    def _working_with_xml(self):

        slides_path = f"{self.source_folder}/ppt/slides"
        self.change_root_pptx_xml()
        self.change_root_pptx_xml_rels()
        i = 0

        for slide in self.slides_to_copy:
            i += 1
            self.update_doc_props(f"{slides_path}/slide{slide}.xml")
            self.change_slide_id(slides_path + '/slide', slide)
            CopyPptxUtils.change_file_index(f"{slides_path}/slide{slide}.xml", i)
            self._change_rels_file(f"{slides_path}/slide{slide}.xml", i, slide)

        self.change_doc_props()
        self.change_root_context_type()
        self.delete_and_move_files(slides_path)

    def update_doc_props(self, slides_path):
        tree = etree.parse(slides_path)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(slides_path)
        a_t = root.findall('.//a:t', namespaces=namespaces)
        for e in a_t:
            self.num_of_words += len(e.text.split())
        self.num_of_paragraphs += len(root.findall('.//a:p', namespaces=namespaces))

    def change_slide_id(self, slides_path, old_index):
        slide_xml_path = f"{slides_path}{old_index}.xml"
        tree = etree.parse(slide_xml_path)
        root = tree.getroot()
        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(slide_xml_path)
        creation_id = root.find('.//p14:creationId', namespaces=namespaces)
        if creation_id is not None:
            creation_id.set('val', f'{random.sample(range(2000000000, 7000000000), 1)[0]}')

        tree.write(slide_xml_path, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_doc_props(self):
        root_pptx_xml = f"{self.source_folder}/docProps/app.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(root_pptx_xml)

        slides = root.find('.//Slides', namespaces=namespaces)
        notes = root.find('.//Notes', namespaces=namespaces)
        paragraphs = root.find('.//Paragraphs', namespaces=namespaces)
        words = root.find('.//Words', namespaces=namespaces)
        current_slides_count = len(self.slides_to_copy)
        paragraphs.text = str(self.num_of_paragraphs)
        words.text = str(self.num_of_words)
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

    def change_root_context_type(self):
        root_pptx_xml = f"{self.source_folder}/[Content_Types].xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces_by_filepath(root_pptx_xml)
        content_type = dict()
        relationship_elements = root.findall('.//Override', namespaces=namespaces)
        relations = None
        for rel in relationship_elements:
            target_type = str(str(rel.get('ContentType')).split('.')[-1].split('+')[0]).lower()

            if target_type == 'slideLayout':
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

        for target_type in self.target_indexes:
            if target_type.lower() not in content_type:
                continue

            for i in range(self.target_indexes[target_type]):
                etree.SubElement(relations, "Override",
                                 {
                                     "PartName": CopyPptxUtils.replace_number(
                                         f"{content_type[target_type.lower()]['pt']}",
                                         str(i + 1)),
                                     "ContentType": f"{content_type[target_type.lower()]['ct']}"
                                 })

        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_root_pptx_xml_rels(self):
        root_pptx_xml = f"{self.source_folder}/ppt/_rels/presentation.xml.rels"
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
            index = self.len_master_id + i + 1
            etree.SubElement(relations, "Relationship",
                             {'Id': f'rId{index}',
                              "Type": f'{all_type}',
                              "Target": f'slides/slide{i + 1}.xml'})

        max_value = max(self.font_ids.values(), key=lambda x: int(x[3:]))
        index = int(str(CopyPptxUtils.get_number_from_str(max_value)[0]))
        for rel in relationship_elements:
            target_type = str(rel.get('Type')).split('/')[-1]
            if target_type == 'slide':
                continue

            if target_type == 'font' \
                    or target_type == 'notesMaster' \
                    or target_type == 'slideMaster':
                id = rel.get("Id")
                if id in self.font_ids:
                    rel.set("Id", f'{self.font_ids[id]}')
            else:
                index += 1
                rel.set("Id", f'rId{index}')
        tree.write(root_pptx_xml, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_root_pptx_xml(self):
        root_pptx_xml = f"{self.source_folder}/ppt/presentation.xml"
        tree = etree.parse(root_pptx_xml)
        root = tree.getroot()

        namespaces = CopyPptxUtils.get_name_spaces(root)
        sld_ids = root.xpath('//ns0:sldId', namespaces=namespaces)
        self.len_master_id = len(root.xpath('//ns0:sldMasterId', namespaces=namespaces))

        for sldId in sld_ids:
            CopyPptxUtils.delete_child(sldId)
        sldIdLst = root.find('ns0:sldIdLst', namespaces=namespaces)
        ids = CopyPptxUtils.generate_ids(len(self.slides_to_copy))
        index = 0
        for i in range(len(self.slides_to_copy)):
            index = i + 1 + self.len_master_id
            etree.SubElement(sldIdLst, "{" + namespaces['ns0'] + "}sldId",
                             {'id': f'{str(ids[i])}',
                              "{" + namespaces['ns1'] + "}id": f'rId{index}'})

        name = "{" + namespaces['ns1'] + "}"
        notes_master_id = root.findall('.//ns0:notesMasterId', namespaces=namespaces)
        for elem in notes_master_id:
            elem_id = elem.get(f'{name}id')
            if elem_id:
                index += 1
                self.font_ids[elem_id] = f'rId{index}'
                elem.set(f"{name}id", f'rId{index}')

        sld_master_id = root.findall('.//ns0:sldMasterId', namespaces=namespaces)
        for elem in sld_master_id:
            elem_id = elem.get(f'{name}id')
            if elem_id:
                self.font_ids[elem_id] = elem_id

        embedded_fonts = root.findall(f'.//ns0:embeddedFont', namespaces=namespaces)
        for embedded_font in embedded_fonts:
            for c in embedded_font.getchildren():
                elem_id = c.get(f'{name}id')
                if elem_id:
                    index += 1
                    self.font_ids[elem_id] = f'rId{index}'
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
        notes_slides_path = f"{self.source_folder}/ppt/notesSlides/notesSlide"
        if old_index in self.repeated_indexes:
            self.repeated_indexes[old_index] += 1

        for rel in relationship_elements:
            target = str(rel.get('Target'))
            target_type = str(rel.get('Type')).split('/')[-1]
            path_to_rel = target.replace('..', self.source_folder + '/ppt')
            index = self.add_target_indexes(target_type)
            if target_type == 'chart':
                CopyPptxUtils.change_chart_id(path_to_rel)
                self._change_chart_rels(path_to_rel, index, old_index)
                rel.set('Target', f'../charts/chart{index}.xml')

            if target_type == 'notesSlide':
                pattern = r'../notesSlides/notesSlide(\d+)\.xml'
                r_num = CopyPptxUtils.extract_slide_numbers(rel.get('Target'), pattern)
                if r_num != index \
                        and r_num is not None:
                    rel.set('Target', f'../notesSlides/notesSlide{index}.xml')
                self.change_slide_id(notes_slides_path, old_index)
                CopyPptxUtils.change_notes_slides(path_to_rel, index)
        tree.write(slide_xml_path_new, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def _change_chart_rels(self, path_to_chart, index, old_index):
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
                self.change_package(rel, chart_target_type, chart_target)
            if chart_target_type == 'chartStyle':
                self.change_chart_style(rel, chart_target_type, chart_target, 'style*', old_index)
            if chart_target_type == 'chartColorStyle':
                self.change_chart_style(rel, chart_target_type, chart_target, 'colors*', old_index)

        tree.write(chart_path_rels, pretty_print=True, xml_declaration=True, encoding='utf-8')

    def change_chart_style(self, rel, chart_target_type, chart_target, pattern, old_index):

        chart_path_to_embedding = self.source_folder + '/ppt/charts'
        embedding_index = str(self.add_target_indexes(chart_target_type))

        CopyPptxUtils.move_file(chart_path_to_embedding + '/' + chart_target, embedding_index)
        rel.set('Target', CopyPptxUtils.replace_number(chart_target, embedding_index))

    def change_package(self, rel, chart_target_type, chart_target):
        embedding_index = str(self.add_target_indexes(chart_target_type))

        chart_path_to_embedding = chart_target.replace('..', self.source_folder + '/ppt')

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
        for root, dirs, files in os.walk(self.source_folder):
            for file in files:
                # print(file)
                source_path = os.path.join(root, file)
                relative_path = os.path.relpath(source_path, self.source_folder)
                try:
                    target_zip.write(source_path, relative_path)
                except OSError as e:
                    logger.warning(e)

    def add_target_indexes(self, target_type):
        if target_type in self.target_indexes:
            self.target_indexes[target_type] = self.target_indexes.get(target_type) + 1
        else:
            self.target_indexes[target_type] = 1

        return self.target_indexes[target_type]

    def get_repeated_indexes(self, numbers):
        counts = {}
        for num in numbers:
            if num in counts:
                counts[num] += 1
            else:
                counts[num] = 1
        rep = [num for num, count in counts.items() if count > 1]
        self.repeated_indexes = dict()
        for i in rep:
            self.repeated_indexes[i] = 0

    def delete_and_move_files(self, slides_path):
        slides_path_note = f"{self.source_folder}/ppt/notesSlides"
        slides_path_charts = f"{self.source_folder}/ppt/charts"
        slides_path_embeddings = f"{self.source_folder}/ppt/embeddings"

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
    source_folder = f"{script_location}/source_pptx_extracted"
    # CopyPptxUtils.search_word_in_xml_folder(source_folder, "Microsoft_Excel_Worksheet")
    # slides_to_copy = random.sample(range(1, 32), 31)
    # slides_to_copy = [i + 1 for i in range(35)]
    pptx_copy = CopyPptx(path_to_source, path_to_new,
                         [1])

    # [23, 17, 28, 8, 26, 30, 22, 19, 2, 21, 9, 29, 14, 12, 15, 13, 5, 24, 10, 25, 18, 4, 11, 16, 20, 1, 6, 31, 27, 7, 3]

    # pptx_copy = CopyPptx(path_to_source, path_to_new,
    #                      [22, 23, 22, 23, 26, 26, 12, 12, 16, 17,
    #                       22, 23, 18, 16, 17, 22, 23, 16, 17, 16,
    #                       17, 32, 16, 17, 18, 18, 22, 23, 26, 26,
    #                       18, 16, 17, 18, 9, 16, 17, 16, 17, 22,
    #                       23, 16, 17, 18, 16, 17, 9, 16, 17, 18,
    #                       22, 23, 18, 9, 9, 18, 16, 17, 22, 23,
    #                       16, 17, 26, 16, 17, 18, 16, 17, 18, 16,
    #                       17, 16, 17, 16, 17, 18, 18, 16, 17, 16,
    #                       17, 18, 16, 17, 16, 17, 16, 17, 26
    #                       ])

    pptx_copy.copy_slides()

    # source_folder = f"{script_location}/res_3"
    # path_to_source = f"{script_location}/res_3.pptx"
    # os.makedirs(source_folder, exist_ok=True)
    #
    # with zipfile.ZipFile(path_to_source, 'r') as source_zip:
    #     source_zip.extractall(source_folder)


main()
