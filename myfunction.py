import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import shutil

def decrypt_readonly_pptx(pptx_file, modified_pptx_file):
    pptx_path = Path(pptx_file)
    modified_pptx_path = Path(modified_pptx_file)
    
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            try:
                xml_content = zip_ref.read('ppt/presentation.xml')
            except KeyError:
                print(f"Warning: 'ppt/presentation.xml' not found in {pptx_path.name}. Copying as is.")
                if pptx_path != modified_pptx_path:
                    shutil.copy2(pptx_path, modified_pptx_path)
                return

            modified_xml = match_attr(xml_content.decode("utf-8"))
            
            with zipfile.ZipFile(modified_pptx_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                for item in zip_ref.infolist():
                    if item.filename != 'ppt/presentation.xml':
                        content = zip_ref.read(item.filename)
                        new_zip.writestr(item, content)
                new_zip.writestr('ppt/presentation.xml', modified_xml)
    except zipfile.BadZipFile:
        print(f"Error: {pptx_path.name} is not a valid zip archive. Copying as is.")
        if pptx_path != modified_pptx_path:
            shutil.copy2(pptx_path, modified_pptx_path)

def match_attr(xml_content: str) -> str:
    try:
        # Register namespaces to prevent ns0 prefixes
        ET.register_namespace('', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        ET.register_namespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')

        root = ET.fromstring(xml_content)
        
        # PPTX uses XML tag like <p:modifyVerifier...> or <modifyVerifier xmlns="..."/>
        element_to_remove = None
        for child in list(root):
            if child.tag.endswith('}modifyVerifier') or child.tag == 'modifyVerifier':
                element_to_remove = child
                break
                
        if element_to_remove is not None:
            root.remove(element_to_remove)
            
        return ET.tostring(root, encoding='utf-8', xml_declaration=True).decode('utf-8')
    except ET.ParseError as e:
        print(f"Warning: XML parsing failed ({e}). Returning original content.")
        return xml_content

def process_directory(source_path, target_path=None, in_place=False):
    src_path = Path(source_path)
    
    if in_place:
        tgt_path = src_path
    else:
        if not target_path:
            raise ValueError("target_path must be provided if not modifying in-place.")
        tgt_path = Path(target_path)

    if src_path.is_file():
        if src_path.suffix.lower() == '.pptx':
            if in_place:
                # Use a temporary file for safe in-place replacement
                temp_file = src_path.with_suffix('.tmp.pptx')
                decrypt_readonly_pptx(src_path, temp_file)
                temp_file.replace(tgt_path)
            else:
                if tgt_path.is_dir():
                    tgt_path.mkdir(parents=True, exist_ok=True)
                    tgt_path = tgt_path / src_path.name
                else:
                    tgt_path.parent.mkdir(parents=True, exist_ok=True)
                decrypt_readonly_pptx(src_path, tgt_path)
    elif src_path.is_dir():
        if not in_place:
            tgt_path.mkdir(parents=True, exist_ok=True)
            
        for source_item_path in src_path.rglob('*'):
            rel_path = source_item_path.relative_to(src_path)
            target_item_path = tgt_path / rel_path
            
            if source_item_path.is_file():
                if source_item_path.suffix.lower() == '.pptx':
                    if in_place:
                        temp_file = source_item_path.with_suffix('.tmp.pptx')
                        decrypt_readonly_pptx(source_item_path, temp_file)
                        temp_file.replace(source_item_path)
                    else:
                        target_item_path.parent.mkdir(parents=True, exist_ok=True)
                        decrypt_readonly_pptx(source_item_path, target_item_path)
                elif not in_place:
                    target_item_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source_item_path, target_item_path)