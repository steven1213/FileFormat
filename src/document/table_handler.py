from docx.shared import Pt
from docx.oxml import parse_xml
import io

class TableHandler:
    @staticmethod
    def copy_table_structure(source_table, target_doc):
        new_table = target_doc.add_table(rows=len(source_table.rows), cols=len(source_table.columns))
        new_table.style = source_table.style
        new_table.alignment = source_table.alignment
        return new_table

    @staticmethod
    def copy_row_height(source_row, target_row):
        if source_row._tr.xpath('./w:trPr/w:trHeight'):
            height = source_row._tr.xpath('./w:trPr/w:trHeight')[0]
            if not target_row._tr.xpath('./w:trPr'):
                target_row._tr.append(parse_xml('<w:trPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
            target_row._tr.trPr.append(height)

    @staticmethod
    def copy_cell_properties(source_cell, target_cell):
        if source_cell._tc.tcPr is not None and source_cell._tc.tcPr.xpath('./w:tcW'):
            target_cell._tc.tcPr.xpath('./w:tcW')[0].attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w'] = \
                source_cell._tc.tcPr.xpath('./w:tcW')[0].attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w']

    @staticmethod
    def copy_image(source_run, target_run):
        for blip in source_run._element.xpath('.//a:blip'):
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rId:
                image_part = source_run.part.rels[rId].target_part
                image_stream = io.BytesIO(image_part.blob)
                inline = source_run._element.xpath('.//wp:inline')[0]
                width = int(inline.extent.cx)
                height = int(inline.extent.cy)
                target_run.add_picture(image_stream, width=width, height=height)