import os
import re
import chardet
import olefile
import io
from PIL import Image

def extract_images_from_doc(doc_path, image_dir):
    try:
        if not olefile.isOleFile(doc_path):
            print(f"File {doc_path} is not a valid OLE file")
            return []
        
        ole = olefile.OleFileIO(doc_path)
        
        if not ole.exists('WordDocument'):
            print(f"No WordDocument stream found in {doc_path}")
            ole.close()
            return []
        
        image_files = []
        
        try:
            data_store = ole.openstream('Data')
            data_store_data = data_store.read()
            print(f"Found Data stream with {len(data_store_data)} bytes")
        except:
            data_store_data = None
        
        for entry in ole.listdir():
            if 'ObjectPool' in entry or 'Pictures' in entry or '1Table' in entry or '0Table' in entry:
                try:
                    stream = ole.openstream(entry)
                    stream_data = stream.read()
                    
                    if len(stream_data) > 1000:
                        try:
                            image = Image.open(io.BytesIO(stream_data))
                            
                            image_name = f"{os.path.basename(doc_path).replace('.doc', '')}_{len(image_files) + 1}.png"
                            image_path = os.path.join(image_dir, image_name)
                            
                            image.save(image_path, 'PNG')
                            image_files.append(image_name)
                            
                            print(f"Extracted image: {image_name} ({image.size[0]}x{image.size[1]})")
                        except Exception as img_error:
                            pass
                except Exception as e:
                    continue
        
        if data_store_data:
            try:
                image_patterns = [
                    b'\xFF\xD8\xFF',
                    b'\x89PNG',
                    b'GIF8',
                    b'BM'
                ]
                
                for pattern in image_patterns:
                    pos = 0
                    while True:
                        pos = data_store_data.find(pattern, pos)
                        if pos == -1:
                            break
                        
                        try:
                            end_pos = len(data_store_data)
                            if pattern == b'\xFF\xD8\xFF':
                                end_marker = data_store_data.find(b'\xFF\xD9', pos)
                                if end_marker != -1:
                                    end_pos = end_marker + 2
                            elif pattern == b'\x89PNG':
                                end_marker = data_store_data.find(b'IEND\xAEB`\x82', pos)
                                if end_marker != -1:
                                    end_pos = end_marker + 8
                            
                            image_data = data_store_data[pos:end_pos]
                            
                            if len(image_data) > 100:
                                try:
                                    image = Image.open(io.BytesIO(image_data))
                                    
                                    image_name = f"{os.path.basename(doc_path).replace('.doc', '')}_{len(image_files) + 1}.png"
                                    image_path = os.path.join(image_dir, image_name)
                                    
                                    image.save(image_path, 'PNG')
                                    image_files.append(image_name)
                                    
                                    print(f"Extracted image from data: {image_name} ({image.size[0]}x{image.size[1]})")
                                except:
                                    pass
                            
                            pos = end_pos
                        except:
                            pos += 1
            except Exception as e:
                print(f"Error scanning for images in data stream: {e}")
        
        ole.close()
        return image_files
        
    except Exception as e:
        print(f"Error extracting images from {doc_path}: {e}")
        return []

def extract_text_from_doc_simple(doc_path):
    try:
        if not olefile.isOleFile(doc_path):
            print(f"File {doc_path} is not a valid OLE file")
            return ""
        
        ole = olefile.OleFileIO(doc_path)
        
        if not ole.exists('WordDocument'):
            print(f"No WordDocument stream found in {doc_path}")
            ole.close()
            return ""
        
        word_doc = ole.openstream('WordDocument')
        word_data = word_doc.read()
        ole.close()
        
        text = ""
        
        encodings = ['utf-16le', 'gb18030', 'gbk', 'gb2312', 'utf-8', 'latin-1', 'cp1252', 'big5', 'shift_jis']
        
        for encoding in encodings:
            try:
                decoded = word_data.decode(encoding, errors='ignore')
                
                # 智能行分割：检测并保留行边界
                lines = []
                current_line = ""
                
                for char in decoded:
                    if char in '\r\n':
                        if current_line.strip():
                            lines.append(current_line.strip())
                        current_line = ""
                    elif ord(char) < 32 or ord(char) == 0xFFFD:
                        continue
                    else:
                        current_line += char
                
                if current_line.strip():
                    lines.append(current_line.strip())
                
                cleaned = '\n'.join(lines)
                
                if len(cleaned) > 100:
                    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', cleaned))
                    total_chars = len(cleaned)
                    ratio = chinese_chars / total_chars if total_chars > 0 else 0
                    
                    if ratio > 0.1 or not re.match(r'^[\s\W]+$', cleaned):
                        text = cleaned
                        print(f"Successfully decoded with {encoding}, Chinese ratio: {ratio:.2f}")
                        break
            except Exception as e:
                continue
        
        if not text or len(text) < 50:
            print("Trying alternative extraction method...")
            
            for i in range(0, len(word_data) - 100, 2):
                try:
                    chunk = word_data[i:i+1000]
                    detected = chardet.detect(chunk)
                    if detected['confidence'] > 0.7:
                        decoded = chunk.decode(detected['encoding'], errors='ignore')
                        if len(decoded) > 50 and re.search(r'[\u4e00-\u9fff]', decoded):
                            text += decoded + "\n"
                except:
                    continue
        
        text = clean_binary_data(text)
        
        return text
    except Exception as e:
        print(f"Error extracting text from {doc_path}: {e}")
        return ""

def clean_binary_data(text):
    if not text:
        return ""
    
    text = re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F-\x9F]', '', text)
    
    patterns_to_remove = [
        r'Root Entry',
        r'SummaryInformation',
        r'DocumentSummaryInformation',
        r'WordDocument',
        r'WPS Office',
        r'Normal\.dotm',
        r'KSOProductBuildVer',
        r'WpsCustomData',
        r'Chow Eee',
        r'F1E327BC-269C-435d-A152-05C5408002CA',
        r'INCLUDEPICTURE.*?MERGEFORMATINET',
        r'HYPERLINK.*?http',
        r'IMG_\d+',
        r'图片\s*\d+',
        r'[\uFFFD\uFFFE\uFFFF]',
    ]
    
    for pattern in patterns_to_remove:
        text = re.sub(pattern, '', text, flags=re.DOTALL)
    
    lines = text.split('\n')
    cleaned_lines = []
    seen_lines = set()
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        if len(line) < 2:
            continue
        
        control_chars = len(re.findall(r'[\x00-\x1F\x7F-\x9F]', line))
        if control_chars > len(line) * 0.4:
            continue
        
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', line))
        ascii_chars = len(re.findall(r'[\x20-\x7E]', line))
        total_chars = len(line)
        
        if total_chars > 0:
            valid_ratio = (chinese_chars + ascii_chars) / total_chars
            if valid_ratio < 0.4:
                continue
        
        if re.match(r'^[^\w\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*]+$', line):
            continue
        
        if line in seen_lines:
            continue
        seen_lines.add(line)
        
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def clean_garbled_text(text):
    if not text:
        return ""
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', line))
        ascii_chars = len(re.findall(r'[\x20-\x7E]', line))
        total_chars = len(line)
        
        if total_chars > 0:
            valid_ratio = (chinese_chars + ascii_chars) / total_chars
            if valid_ratio < 0.3:
                continue
        
        other_unicode = len(re.findall(r'[^\x20-\x7E\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*\n\r]', line))
        if other_unicode > len(line) * 0.3:
            line = re.sub(r'[^\x20-\x7E\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*\n\r]+', '', line)
        
        if line:
            cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def clean_final_text(text):
    if not text:
        return ""
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', line))
        ascii_chars = len(re.findall(r'[\x20-\x7E]', line))
        total_chars = len(line)
        
        if total_chars > 0:
            valid_ratio = (chinese_chars + ascii_chars) / total_chars
            if valid_ratio < 0.4:
                continue
        
        special_chars = len(re.findall(r'[^\x20-\x7E\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*\n\r]', line))
        if special_chars > len(line) * 0.2:
            line = re.sub(r'[^\x20-\x7E\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*\n\r]+', '', line)
            if len(line.strip()) < 2:
                continue
        
        if re.match(r'^[^\w\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*]+$', line):
            continue
        
        if len(line) < 2:
            continue
        
        cleaned_lines.append(line)
    
    result = '\n'.join(cleaned_lines)
    
    result = re.sub(r'^[^\w\u4e00-\u9fff\s.,;:!?()\-+=\[\]{}"\'/\\@#$%^&*]+', '', result, flags=re.MULTILINE)
    
    return result

def clean_heading_lines(text):
    if not text:
        return ""
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        
        if line.startswith('#'):
            heading_match = re.match(r'^(#+)\s*(.*)$', line)
            if heading_match:
                heading_level = heading_match.group(1)
                heading_content = heading_match.group(2).strip()
                
                chinese_matches = re.findall(r'[\u4e00-\u9fff]{2,}', heading_content)
                if chinese_matches:
                    cleaned_content = ' '.join(chinese_matches)
                    cleaned_line = f"{heading_level} {cleaned_content}"
                    cleaned_lines.append(cleaned_line)
                else:
                    ascii_matches = re.findall(r'[\x20-\x7E]{3,}', heading_content)
                    if ascii_matches:
                        cleaned_content = ' '.join(ascii_matches)
                        cleaned_line = f"{heading_level} {cleaned_content}"
                        cleaned_lines.append(cleaned_line)
                    else:
                        cleaned_lines.append(line)
            else:
                cleaned_lines.append(line)
        else:
            cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def clean_all_lines(text):
    if not text:
        return ""
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        
        chinese_matches = re.findall(r'[\u4e00-\u9fff]{2,}', line)
        if chinese_matches:
            cleaned_content = ' '.join(chinese_matches)
            cleaned_lines.append(cleaned_content)
        else:
            ascii_matches = re.findall(r'[\x20-\x7E]{3,}', line)
            if ascii_matches:
                cleaned_content = ' '.join(ascii_matches)
                cleaned_lines.append(cleaned_content)
            else:
                cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def clean_garbled_text_advanced(text):
    if not text:
        return ""
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        
        chinese_matches = re.findall(r'[\u4e00-\u9fff]{2,}', line)
        if chinese_matches:
            cleaned_content = ' '.join(chinese_matches)
            cleaned_lines.append(cleaned_content)
        else:
            ascii_matches = re.findall(r'[\x20-\x7E]{3,}', line)
            if ascii_matches:
                cleaned_content = ' '.join(ascii_matches)
                cleaned_lines.append(cleaned_content)
            else:
                cleaned_lines.append("")
    
    return '\n'.join(cleaned_lines)

def clean_uncommon_chinese(text):
    if not text:
        return ""
    
    uncommon_chars = set('袉倔卋卋尀伀倀儀帀漀伨伥昀焁洀猄渄琈弈愀')
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        
        filtered_line = ""
        for char in line:
            if char not in uncommon_chars:
                filtered_line += char
        
        if filtered_line.strip():
            chinese_matches = re.findall(r'[\u4e00-\u9fff]{2,}', filtered_line)
            if chinese_matches:
                cleaned_content = ' '.join(chinese_matches)
                cleaned_lines.append(cleaned_content)
            else:
                ascii_matches = re.findall(r'[\x20-\x7E]{3,}', filtered_line)
                if ascii_matches:
                    cleaned_content = ' '.join(ascii_matches)
                    cleaned_lines.append(cleaned_content)
                else:
                    cleaned_lines.append(filtered_line)
        else:
            cleaned_lines.append("")
    
    return '\n'.join(cleaned_lines)

def format_markdown(text, title):
    lines = text.split('\n')
    formatted_lines = []
    
    formatted_lines.append(f"# {title}\n")
    
    for line in lines:
        line = line.strip()
        if not line:
            formatted_lines.append("")
            continue
        
        if re.match(r'^[一二三四五六七八九十]+[、.]\s*', line):
            line = re.sub(r'^([一二三四五六七八九十]+)[、.]\s*', r'## \1、', line)
        elif re.match(r'^\d+[、.]\s*', line):
            line = re.sub(r'^(\d+)[、.]\s*', r'## \1. ', line)
        elif re.match(r'^[A-Z][A-Z\s]+：', line):
            line = re.sub(r'^([A-Z][A-Z\s]+)：', r'## \1：', line)
        elif re.match(r'^【.*】', line):
            line = re.sub(r'^【(.*)】', r'## \1', line)
        
        if re.search(r'\b(site|inurl|intitle|intext|filetype|link|cache|info|define)\b', line, re.IGNORECASE):
            if not line.startswith('#'):
                line = f"## {line}"
        
        formatted_lines.append(line)
    
    return '\n'.join(formatted_lines)

def process_doc_files():
    current_dir = os.getcwd()
    image_dir = r'D:\0-MD图床'
    
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)
        print(f"Created image directory: {image_dir}")
    
    doc_files = [f for f in os.listdir(current_dir) if f.endswith('.doc') and not f.startswith('~$')]
    
    for doc_file in doc_files:
        doc_path = os.path.join(current_dir, doc_file)
        md_path = os.path.join(current_dir, doc_file.replace('.doc', '.md'))
        
        print(f"Processing: {doc_path}")
        
        print("Extracting images...")
        image_files = extract_images_from_doc(doc_path, image_dir)
        print(f"Found {len(image_files)} images")
        
        text = extract_text_from_doc_simple(doc_path)
        
        if text and len(text) > 50:
            text = clean_garbled_text(text)
            text = clean_final_text(text)
            text = clean_garbled_text_advanced(text)
            text = clean_uncommon_chinese(text)
            text = clean_all_lines(text)
            
            title = doc_file.replace('.doc', '')
            formatted_text = format_markdown(text, title)
            formatted_text = clean_heading_lines(formatted_text)
            
            if image_files:
                image_section = "\n\n## 图片\n\n"
                for i, image_file in enumerate(image_files, 1):
                    image_section += f"![图片{i}](../0-MD图床/{image_file})\n\n"
                formatted_text += image_section
            
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(formatted_text)
            
            print(f"Converted: {doc_path} -> {md_path}")
            print(f"Extracted {len(text)} characters")
            print(f"Added {len(image_files)} images")
        else:
            print(f"Failed to extract sufficient text from {doc_path}")
        
        print()

if __name__ == "__main__":
    process_doc_files()
