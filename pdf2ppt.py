import logging
import os
import requests
import tempfile
import time
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
import shutil
import uuid
import json
import zipfile
import io
import hashlib
import math
from pathlib import Path

# --- Import Configuration ---
try:
    import config
except ImportError:
    # Fallback if config.py is missing
    logging.warning("config.py not found, using default settings.")
    class ConfigMock:
        MINERU_TOKEN = ""
        PDF_INPUT_PATH = "input.pdf"
        PPT_OUTPUT_PATH = "output.pptx"
        CACHE_DIR = "temp"
        USE_CACHE = True
        PPT_SLIDE_WIDTH = 16
        PPT_SLIDE_HEIGHT = 9
    config = ConfigMock()

# --- Helpers ---

def get_pdf_hash(pdf_path):
    """计算 PDF 文件的 MD5 哈希（用于缓存标识）"""
    hasher = hashlib.md5()
    with open(pdf_path, 'rb') as f:
        hasher.update(f.read())
    return hasher.hexdigest()[:8]

class MinerUClient:
    def __init__(self, token):
        self.token = token
        self.base_url = "https://mineru.net/api/v4"
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

    def upload_and_extract(self, file_path, model_version="vlm"):
        """上传 PDF 并等待自动解析"""
        file_name = os.path.basename(file_path)
        data_id = str(uuid.uuid4())
        
        logging.info(f"为文件 {file_name} 申请上传 URL...")
        batch_url = f"{self.base_url}/file-urls/batch"
        data = {
            "files": [{"name": file_name, "data_id": data_id}],
            "model_version": model_version
        }
        
        response = requests.post(batch_url, headers=self.headers, json=data)
        response.raise_for_status()
        result = response.json()
        
        if result.get("code") != 0:
            raise Exception(f"申请上传 URL 失败: {result.get('msg')}")
        
        batch_id = result["data"]["batch_id"]
        upload_url = result["data"]["file_urls"][0]
        
        logging.info(f"✓ 获得 Batch ID: {batch_id}")
        
        # 上传文件
        with open(file_path, 'rb') as f:
            upload_response = requests.put(upload_url, data=f, timeout=120)
            upload_response.raise_for_status()
            
        logging.info(f"✓ 文件上传成功")
        return batch_id

    def get_batch_result(self, batch_id):
        """轮询任务结果并解析 ZIP"""
        result_url = f"{self.base_url}/extract-results/batch/{batch_id}"
        logging.info(f"轮询中...")
        
        poll_count = 0
        max_polls = 180
        
        while poll_count < max_polls:
            response = requests.get(result_url, headers={"Authorization": f"Bearer {self.token}"})
            response.raise_for_status()
            result = response.json()
            
            if result.get("code") != 0:
                raise Exception(f"查询失败: {result.get('msg')}")
            
            extract_results = result.get("data", {}).get("extract_result", [])
            if not extract_results:
                raise Exception("未找到解析结果")
            
            file_result = extract_results[0]
            state = file_result.get("state")
            
            poll_count += 1
            
            if state == "done":
                full_zip_url = file_result.get("full_zip_url")
                if not full_zip_url:
                    raise Exception("未返回结果 URL")
                
                logging.info(f"✓ 解析完成，下载中...")
                zip_response = requests.get(full_zip_url, timeout=120)
                zip_response.raise_for_status()
                
                return zip_response.content
                
            elif state == "failed":
                raise Exception(f"解析失败: {file_result.get('err_msg')}")
            
            if poll_count % 3 == 1:  # 每30秒显示一次
                logging.info(f"  状态: {state} ({poll_count}/{max_polls})")
            time.sleep(10)
        
        raise Exception(f"任务超时")

def ensure_cache_dir(cache_path):
    """确保缓存目录存在"""
    if not os.path.exists(cache_path):
        os.makedirs(cache_path)
        logging.info(f"创建缓存目录: {cache_path}")

def save_mineru_result_to_cache(page_num, zip_content, pdf_hash, cache_dir):
    """保存 MinerU 结果到缓存"""
    ensure_cache_dir(cache_dir)
    
    page_dir = os.path.join(cache_dir, f"{pdf_hash}_page_{page_num}")
    if not os.path.exists(page_dir):
        os.makedirs(page_dir)
    
    # 1. 保存 ZIP
    zip_path = os.path.join(page_dir, "mineru_result.zip")
    with open(zip_path, 'wb') as f:
        f.write(zip_content)
    logging.info(f"  ✓ 保存 ZIP: {zip_path}")
    
    # 2. 解析并保存 JSON 和图片
    zip_data = io.BytesIO(zip_content)
    with zipfile.ZipFile(zip_data, 'r') as zip_ref:
        # 保存所有 JSON 文件
        for file_name in zip_ref.namelist():
            if file_name.endswith('.json'):
                json_content = zip_ref.read(file_name)
                json_path = os.path.join(page_dir, os.path.basename(file_name))
                with open(json_path, 'wb') as f:
                    f.write(json_content)
                logging.info(f"  ✓ 保存 JSON: {json_path}")
        
        # 保存图片  
        images_dir = os.path.join(page_dir, "images")
        if not os.path.exists(images_dir):
            os.makedirs(images_dir)
        
        for file_name in zip_ref.namelist():
            if file_name.startswith('images/') and not file_name.endswith('/'):
                img_data = zip_ref.read(file_name)
                img_path = os.path.join(images_dir, os.path.basename(file_name))
                with open(img_path, 'wb') as f:
                    f.write(img_data)
                logging.info(f"  ✓ 保存图片: {img_path}")
    
    logging.info(f"✓ 缓存已保存到: {page_dir}")

def load_mineru_result_from_cache(page_num, pdf_hash, cache_dir):
    """从缓存加载 MinerU 结果"""
    page_dir = os.path.join(cache_dir, f"{pdf_hash}_page_{page_num}")
    zip_path = os.path.join(page_dir, "mineru_result.zip")
    
    if os.path.exists(zip_path):
        logging.info(f"  ✓ 从缓存加载: {page_dir}")
        with open(zip_path, 'rb') as f:
            return f.read()
    
    return None

def recursive_blocks(blocks):
    """递归提取所有 Layout Block (借鉴 MinerU 优化思路)"""
    result = []
    for block in blocks:
        if isinstance(block, dict):
            # 如果 block 自身包含 blocks (嵌套布局)
            if "blocks" in block:
                result.extend(recursive_blocks(block["blocks"]))
            else:
                result.append(block)
    return result

def parse_mineru_zip(zip_content):
    """解析 MinerU ZIP（使用 content_list.json）"""
    zip_data = io.BytesIO(zip_content)
    result = {
        'elements': [],
        'images_data': {},
        'page_size': None
    }
    
    with zipfile.ZipFile(zip_data, 'r') as zip_ref:
        # 1. 优先读取 content_list.json
        content_list_files = [f for f in zip_ref.namelist() 
                             if f.endswith('_content_list.json')]
        
        if content_list_files:
            json_content = zip_ref.read(content_list_files[0])
            raw_elements = json.loads(json_content)
            # 使用递归展平逻辑，确保存储在 content_list 中的嵌套结构也能被读取
            result['elements'] = recursive_blocks(raw_elements)
            logging.debug(f"读取 {len(result['elements'])} 个元素 (Recursive)")
        
        # 2. 提取图片
        image_files = [f for f in zip_ref.namelist() 
                      if f.startswith('images/') and not f.endswith('/')]
        
        for img_path in image_files:
            img_name = os.path.basename(img_path)
            result['images_data'][img_name] = zip_ref.read(img_path)
    
    return result

def calculate_font_size_gemini_style(block_fontSize, page_height, slide_height_inches=7.5):
    """
    借鉴 Gemini Canvas 版本的字体计算公式
    """
    # 72 pt per inch
    scale_factor = (slide_height_inches * 72) / page_height
    font_size = block_fontSize * scale_factor * 0.8
    
    # 限制范围
    font_size = max(6, min(72, font_size))
    
    return int(font_size)

def estimate_font_size_by_area(bbox_w, bbox_h, char_count, is_title=False):
    """根据面积和字符数估算字号"""
    if char_count <= 0:
        return 14
        
    # 如果是标题(字数少)，可以直接用高度估算
    if char_count < 15 and bbox_w / bbox_h > 2:
        return bbox_h * 0.7
    
    # 面积法公式：FontSize = sqrt(Area / (K * CharCount))
    # K 是单个字符占用的平均面积系数，包含行距等
    # 0.8 是经验值
    area = bbox_w * bbox_h
    estimated_fontSize_px = math.sqrt(area / (0.8 * char_count))
    
    # 限制字号不能超过 bbox 高度的一定比例 (防止单行时溢出)
    estimated_fontSize_px = min(estimated_fontSize_px, bbox_h * 0.9)
    
    return estimated_fontSize_px

def add_image_to_slide(slide, img_path, ppt_x, ppt_y, ppt_width, ppt_height, images_data, temp_dir):
    """辅助函数：添加图片到幻灯片"""
    img_filename = os.path.basename(img_path)
    if img_filename in images_data:
        temp_img_path = os.path.join(temp_dir, img_filename)
        try:
            with open(temp_img_path, 'wb') as f:
                f.write(images_data[img_filename])
            
            slide.shapes.add_picture(
                temp_img_path, 
                ppt_x, 
                ppt_y, 
                width=ppt_width,
                height=ppt_height
            )
            return True
        except Exception as e:
            logging.warning(f"  图片添加失败: {img_filename}, {e}")
            return False
        finally:
            if os.path.exists(temp_img_path):
                try:
                    os.unlink(temp_img_path)
                except:
                    pass
    return False

def process_elements(elements, slide, images_data, page_width, page_height, 
                       slide_width_emu, slide_height_emu, slide_height_inches, temp_dir):
    """
    处理所有元素并添加到 Slide
    """
    stats = {'text': 0, 'title': 0, 'footer': 0, 'image': 0, 'table': 0, 'list': 0, 'other': 0, 'skipped': 0}
    
    for i, element in enumerate(elements):
        if not isinstance(element, dict):
            stats['skipped'] += 1
            continue
        
        elem_type = element.get('type')
        bbox = element.get('bbox')
        
        if not bbox or len(bbox) != 4:
            logging.debug(f"  元素 {i}: type={elem_type}, bbox=无效")
            stats['skipped'] += 1
            continue
        
        # content_list.json 的 bbox 是像素坐标
        x1_px, y1_px, x2_px, y2_px = bbox
        
        # 转换为 PPT EMU 坐标
        x1_ratio = x1_px / page_width
        y1_ratio = y1_px / page_height
        x2_ratio = x2_px / page_width
        y2_ratio = y2_px / page_height
        
        ppt_x = Emu(int(x1_ratio * slide_width_emu))
        ppt_y = Emu(int(y1_ratio * slide_height_emu))
        ppt_width = Emu(int((x2_ratio - x1_ratio) * slide_width_emu))
        ppt_height = Emu(int((y2_ratio - y1_ratio) * slide_height_emu))
        
        # 最小尺寸保护
        min_width = Emu(int(0.01 * slide_width_emu))
        min_height = Emu(int(0.01 * slide_height_emu))
        ppt_width = max(ppt_width, min_width)
        ppt_height = max(ppt_height, min_height)
        
        # 1. 处理图片
        if elem_type == 'image':
            img_path = element.get('img_path')
            if img_path:
                if add_image_to_slide(slide, img_path, ppt_x, ppt_y, ppt_width, ppt_height, images_data, temp_dir):
                    stats['image'] += 1
                else:
                    stats['skipped'] += 1
            else:
                stats['skipped'] += 1

        # 2. 处理表格 (优先使用图片)
        elif elem_type == 'table':
            img_path = element.get('img_path')
            if img_path and add_image_to_slide(slide, img_path, ppt_x, ppt_y, ppt_width, ppt_height, images_data, temp_dir):
                stats['table'] += 1
                logging.debug(f"  元素 {i}: table (as image)")
            else:
                # 降级处理：作为文本
                table_text = element.get('text') or element.get('content', '')
                if not table_text and element.get('html'):
                    table_text = "[表格数据]"
                
                if table_text:
                    try:
                        txBox = slide.shapes.add_textbox(ppt_x, ppt_y, ppt_width, ppt_height)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = table_text
                        p.font.size = Pt(10)
                        p.font.name = 'Consolas'
                        stats['table'] += 1
                    except:
                        stats['skipped'] += 1
                else:
                    stats['skipped'] += 1
        
        # 3. 处理文字 (Text/Title/Footer) & 列表 (List)
        elif elem_type in ['text', 'title', 'footer', 'list']:
            text_content = ""
            is_title = (elem_type == 'title')
            
            # 提取文本内容
            if elem_type == 'list':
                list_items = element.get('list_items') or []
                if list_items:
                    # 将列表项合并，并在每项前加点
                    text_content = "\n".join([f"{item}" for item in list_items])
            else:
                text_content = element.get('text') or element.get('content', '')

            if not text_content:
                stats['skipped'] += 1
                continue
            
            try:
                txBox = slide.shapes.add_textbox(ppt_x, ppt_y, ppt_width, ppt_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = text_content
                
                # --- 字号计算 (核心优化) ---
                bbox_w = x2_px - x1_px
                bbox_h = y2_px - y1_px
                char_count = len(text_content) # 不去除换行符，因为换行也占空间
                
                estimated_fontSize_px = estimate_font_size_by_area(bbox_w, bbox_h, char_count, is_title)
                
                # 转换单位
                font_size = calculate_font_size_gemini_style(
                    estimated_fontSize_px, 
                    page_height, 
                    slide_height_inches
                )
                
                # 类型微调
                if elem_type == 'title':
                    font_size = max(font_size, 20)
                    p.font.bold = True
                elif elem_type == 'footer':
                    font_size = min(font_size, 10)
                elif elem_type == 'list':
                    # 列表字号稍微缩小一点，防止过于拥挤
                    font_size = min(font_size, 20)
                
                p.font.size = Pt(font_size)
                p.font.name = 'Microsoft YaHei'
                
                stats[elem_type] += 1
                logging.debug(f"  元素 {i}: {elem_type}, len={char_count}, size={font_size}")
            except Exception as e:
                logging.warning(f"  文本添加失败: {e}")
                stats['skipped'] += 1
        
        else:
            stats['other'] += 1

    logging.info(f"  统计: {stats}")

def convert_pdf_to_ppt(pdf_input_path, ppt_output_path, mineru_token, 
                         ppt_slide_width=16, ppt_slide_height=9, 
                         use_cache=True, cache_dir="temp"):
    """
    核心转换函数
    """
    if not os.path.exists(pdf_input_path):
        logging.error(f"输入 PDF 不存在: {pdf_input_path}")
        raise FileNotFoundError(f"文件不存在: {pdf_input_path}")
    
    if not mineru_token:
        logging.error("MINERU_TOKEN 未设置")
        raise ValueError("Token不能为空")

    pdf_hash = get_pdf_hash(pdf_input_path)
    logging.info(f"PDF 哈希: {pdf_hash}")

    temp_dir = tempfile.mkdtemp()
    
    try:
        client = MinerUClient(mineru_token)
        tasks = []
        
        logging.info(f"正在拆分 PDF '{pdf_input_path}'...")
        doc = fitz.open(pdf_input_path)
        
        # --- 第一阶段：批量提交任务 (Phase 1: Batch Submission) ---
        logging.info(f"正在批量上传 {len(doc)} 个页面以并发处理...")
        
        for i, page in enumerate(doc):
            page_num = i + 1
            
            # Check Cache
            cached_zip = None
            if use_cache:
                cached_zip = load_mineru_result_from_cache(page_num, pdf_hash, cache_dir)
            
            if cached_zip:
                logging.info(f"  [页 {page_num}] 命中本地缓存")
                tasks.append({
                    'type': 'cached',
                    'page_num': page_num,
                    'content': cached_zip
                })
            else:
                # Prepare single page PDF
                single_page_doc = fitz.open()
                single_page_doc.insert_pdf(doc, from_page=i, to_page=i)
                page_path = os.path.join(temp_dir, f"page_{page_num}.pdf")
                single_page_doc.save(page_path)
                
                # Upload
                try:
                    batch_id = client.upload_and_extract(page_path)
                    tasks.append({
                        'type': 'api',
                        'page_num': page_num,
                        'batch_id': batch_id
                    })
                except Exception as e:
                    logging.error(f"  [页 {page_num}] 上传失败: {e}")
                    # 如果上传失败，记录错误，后续处理时忽略或报错
                    tasks.append({
                        'type': 'error',
                        'page_num': page_num,
                        'error': str(e)
                    })

        doc.close() # Close original doc to free file handle

        # --- 第二阶段：获取结果 (Phase 2: Result Retrieval) ---
        # 注意：这里我们按顺序获取结果。
        # 虽然是顺序等待，但由于所有任务都已经在服务器排队，
        # 当我们在等第1页时，第2-N页也在后台处理中。
        # 这样就实现了服务端的“并发”。
        
        all_pages_results = []
        
        logging.info(f"\n{'-'*40}\n所有页面已提交，开始轮询结果\n{'-'*40}")
        
        for task in tasks:
            page_num = task['page_num']
            
            if task['type'] == 'error':
                logging.warning(f"[页 {page_num}] 跳过 (上传阶段失败)")
                all_pages_results.append(None)
                continue
                
            if task['type'] == 'cached':
                zip_content = task['content']
                logging.info(f"[页 {page_num}] 使用缓存数据")
            else:
                batch_id = task['batch_id']
                logging.info(f"[页 {page_num}] 正在等待任务 {batch_id} ...")
                try:
                    zip_content = client.get_batch_result(batch_id)
                    if use_cache:
                        save_mineru_result_to_cache(page_num, zip_content, pdf_hash, cache_dir)
                except Exception as e:
                    logging.error(f"[页 {page_num}] 解析失败: {e}")
                    all_pages_results.append(None)
                    continue

            # Parse Result
            page_result = parse_mineru_zip(zip_content)
            
            if page_result and page_result.get('elements'):
                # 推测页面尺寸
                max_x = max_y = 0
                for elem in page_result['elements']:
                    if elem.get('bbox') and len(elem['bbox']) == 4:
                        max_x = max(max_x, elem['bbox'][2])
                        max_y = max(max_y, elem['bbox'][3])
                
                page_width = max_x if max_x > 0 else 1000
                page_height = max_y if max_y > 0 else 1000
                
                all_pages_results.append({
                    'elements': page_result['elements'],
                    'images_data': page_result.get('images_data', {}),
                    'page_width': page_width,
                    'page_height': page_height
                })
            else:
                logging.warning(f"[页 {page_num}] 未返回有效内容")
                all_pages_results.append(None)

        # --- 第三阶段：生成 PPT (Phase 3: Generation) ---
        logging.info(f"\n{'='*40}\n开始合成 PPT\n{'='*40}")
        prs = Presentation()
        prs.slide_width = Inches(ppt_slide_width)
        prs.slide_height = Inches(ppt_slide_height)
        slide_width_emu = prs.slide_width
        slide_height_emu = prs.slide_height

        for i, page_data in enumerate(all_pages_results):
            logging.info(f"生成幻灯片 {i+1}...")
            # 6 是空白布局 (Blank)
            slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(slide_layout)
            
            if page_data:
                process_elements(
                    page_data['elements'],
                    slide,
                    page_data['images_data'],
                    page_data['page_width'],
                    page_data['page_height'],
                    slide_width_emu,
                    slide_height_emu,
                    ppt_slide_height,
                    temp_dir
                )
            else:
                # 添加一个错误提示文本框
                txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
                txBox.text_frame.text = f"Page {i+1} 解析失败或无内容"

        prs.save(ppt_output_path)
        logging.info(f"\n✅ 完成: {ppt_output_path}")
        return ppt_output_path

    except Exception as e:
        logging.error(f"失败: {e}", exc_info=True)
        raise e
    finally:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

def main():
    # 配置日志 (CLI 模式下默认输出到 debug.log 和 控制台)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("debug.log", encoding='utf-8', mode='w'),
            logging.StreamHandler()
        ]
    )

    if not config.PDF_INPUT_PATH:
        logging.error("配置文件中 PDF_INPUT_PATH 为空")
        return

    convert_pdf_to_ppt(
        pdf_input_path=config.PDF_INPUT_PATH,
        ppt_output_path=config.PPT_OUTPUT_PATH,
        mineru_token=config.MINERU_TOKEN,
        ppt_slide_width=config.PPT_SLIDE_WIDTH,
        ppt_slide_height=config.PPT_SLIDE_HEIGHT,
        use_cache=config.USE_CACHE,
        cache_dir=config.CACHE_DIR
    )

if __name__ == "__main__":
    main()
