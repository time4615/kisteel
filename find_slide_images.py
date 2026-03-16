import zipfile
import re
import xml.etree.ElementTree as ET

pptx_path = "(주)케이아이스틸 회사소개서 2025ver.pptx"

with open('slide_images_utf8.txt', 'w', encoding='utf-8') as f:
    with zipfile.ZipFile(pptx_path, 'r') as z:
        for i in range(1, 25): # assuming max 24 slides
            rels_path = f"ppt/slides/_rels/slide{i}.xml.rels"
            if rels_path in z.namelist():
                data = z.read(rels_path).decode('utf-8')
                root = ET.fromstring(data)
                images = []
                for child in root:
                    if 'image' in child.attrib.get('Target', ''):
                        images.append(child.attrib['Target'])
                if images:
                    f.write(f"Slide {i} uses images: {', '.join(images)}\n")
