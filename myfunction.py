import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
def decrypt_readonly_pptx(pptx_file,modified_pptx_file):
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
        xml_content = zip_ref.read('ppt/presentation.xml')
        modified_xml = match_attr(xml_content.decode("utf-8"))
        with zipfile.ZipFile(modified_pptx_file, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            for item in zip_ref.infolist():
                if item.filename != 'ppt/presentation.xml':
                    content = zip_ref.read(item.filename)
                    new_zip.writestr(item, content)
            new_zip.writestr('ppt/presentation.xml', modified_xml)
# 指定 .pptx 文件路径

def match_attr(xml_content:str)->str:
    # 查找 "modifyVerifier" 的起始位置
    start_index = xml_content.find("modifyVerifier")
    
    tag_start_index = xml_content.rfind("<", 0, start_index)  # 查找 "<" 的位置
    tag_end_index = xml_content.find(">", start_index)  # 查找 ">" 的位置

    ret_content = xml_content[:tag_start_index] + xml_content[tag_end_index+1:]
    return ret_content

def process_directory(source_path, target_path):
    if os.path.isfile(source_path):
        # 如果源路径是文件，检查是否为 .pptx 文件
        if source_path.endswith('.pptx') and target_path.endswith('.pptx'):
            decrypt_readonly_pptx(source_path, target_path)
    else:
        os.makedirs(target_path, exist_ok=True)
        
        # 初始化堆栈
        stack = [(source_path, target_path)]
        while stack:
            source_path, target_path = stack.pop()
            
            # 遍历源目录中的所有文件和子目录
            for item_name in os.listdir(source_path):
                source_item_path = os.path.join(source_path, item_name)
                target_item_path = os.path.join(target_path, item_name)
                
                if os.path.isfile(source_item_path) and source_item_path.endswith('.pptx'):
                    decrypt_readonly_pptx(source_item_path, target_item_path)
                elif os.path.isfile(source_item_path):
                    shutil.copy2(source_item_path, target_item_path)
                elif os.path.isdir(source_item_path):
                    os.makedirs(target_item_path, exist_ok=True)
                    stack.append((source_item_path, target_item_path))
def main():
    pptx_file = R"C:/Users/80163/Desktop/4.2 ppt.pptx"
    modified_pptx_file = R'C:/Users/80163/Desktop/temp.pptx'
    decrypt_readonly_pptx(pptx_file)

if __name__ == "__main__":
    main()