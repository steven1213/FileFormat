import io
import logging
from docx.shared import Pt
from PIL import Image as PILImage

def process_image(run, new_paragraph, config):
    """Process and add images from the source document to the new paragraph."""
    for blip in run._element.xpath('.//a:blip'):
        try:
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rId:
                image_part = run.part.rels[rId].target_part
                image_stream = io.BytesIO(image_part.blob)
                
                # 获取图片尺寸
                img = PILImage.open(image_stream)
                width, height = img.size
                
                # 重置流位置
                image_stream.seek(0)
                
                # 添加图片到新段落
                new_paragraph.add_run().add_picture(image_stream, width=Pt(width), height=Pt(height))
                
        except Exception as e:
            logging.error(f'处理图片时出错: {str(e)}')
            continue

def resize_image(image_stream, max_width=None, max_height=None):
    """调整图片大小，保持宽高比"""
    img = PILImage.open(image_stream)
    width, height = img.size
    
    if max_width and width > max_width:
        ratio = max_width / width
        width = max_width
        height = int(height * ratio)
    
    if max_height and height > max_height:
        ratio = max_height / height
        height = max_height
        width = int(width * ratio)
    
    if width != img.size[0] or height != img.size[1]:
        img = img.resize((width, height), PILImage.LANCZOS)
        
    output = io.BytesIO()
    img.save(output, format=img.format)
    output.seek(0)
    return output, width, height