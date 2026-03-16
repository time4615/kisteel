import zipfile
import os
import shutil

pptx_path = "(주)케이아이스틸 회사소개서 2025ver.pptx"
output_dir = "ppt_media"

if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir)

with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
    for file_info in zip_ref.infolist():
        if file_info.filename.startswith("ppt/media/"):
            zip_ref.extract(file_info, output_dir)
            print(f"Extracted: {file_info.filename}")

print("Media extraction complete.")
