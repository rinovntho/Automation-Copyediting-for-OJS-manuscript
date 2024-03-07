import os
import shutil
import docx
import zipfile
def main():
    file_name_j = os.listdir('journal/')[0]
    file_path_j = os.path.join('journal/', file_name_j)
    # journal_doc = docx.Document(file_path_j)

    file_name_o = os.listdir('output/')[0]
    file_path_o = os.path.join('output/', file_name_o)
    # output_docx = docx.Document(file_name_o)

    switch_media(file_path_j, file_path_o)


def switch_media(source_path, target_path):
    base_name_src, extension = os.path.splitext(source_path)
    new_src_path = f"{base_name_src}.zip"
    os.rename(source_path, new_src_path)

    with zipfile.ZipFile(new_src_path, 'r') as file:
        file.extractall(base_name_src)

    base_name_dst, extension = os.path.splitext(target_path)
    new_dst_path = f"{base_name_dst}.zip"
    os.rename(target_path, new_dst_path)

    with zipfile.ZipFile(new_dst_path, 'r') as file:
        file.extractall(base_name_dst)

    src_media_path = f'{base_name_src}/word/media'
    dst_media_path = f'{base_name_dst}/word/media'

    for img in os.listdir(dst_media_path):
        print("IMAGE:", img)
        img_path = os.path.join(dst_media_path, img)
        if os.path.isfile(img_path):
            os.remove(img_path)

    for img in os.listdir(src_media_path):
        print("IMAGE:",img)
        img_path = os.path.join(src_media_path, img)
        dst_path = os.path.join(dst_media_path, img)
        shutil.copy(img_path, dst_path)

    with zipfile.ZipFile(f'{base_name_src}.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(base_name_src):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, src_media_path))

    with zipfile.ZipFile(f'{base_name_dst}.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(base_name_dst):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, dst_media_path))

    os.rename(new_src_path, f'{base_name_src}.docx')
    os.rename(new_dst_path, f'{base_name_dst}.docx')

# Usage example
main()
