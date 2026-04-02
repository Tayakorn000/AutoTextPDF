import fitz # PyMuPDF
import pandas as pd
import os
import re
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageEnhance
import textwrap
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import unicodedata
try:
    from pythainlp.util import normalize as thai_normalize
except ImportError:
    thai_normalize = lambda x: x

import shutil
import sys

class PDFProcessor:
    def __init__(self, db_name='products_db.xlsx'):
        self.db_name = db_name
        self.db_path = self._get_persistent_db_path()
        self._load_db()
        if os.name == 'nt':
            self.font_path = "C:\\Windows\\Fonts\\tahoma.ttf"
            if not os.path.exists(self.font_path): self.font_path = "C:\\Windows\\Fonts\\arial.ttf"
        else:
            self.font_path = "/Library/Fonts/Thonburi.ttc"
            if not os.path.exists(self.font_path): self.font_path = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
        
        if not self.font_path or not os.path.exists(self.font_path):
            self.font_path = None

    def _get_persistent_db_path(self):
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            bundle_dir = sys._MEIPASS
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
            bundle_dir = exe_dir

        local_db = os.path.join(exe_dir, self.db_name)
        bundled_db = os.path.join(bundle_dir, self.db_name)

        if not os.path.exists(local_db) and os.path.exists(bundled_db):
            try: shutil.copy2(bundled_db, local_db)
            except: pass
        
        return local_db

    def _normalize_for_match(self, text):
        if pd.isna(text) or text is None: return ""
        text = str(text)
        if text.lower() == 'nan': return ""
        text = text.lower().strip()
        text = re.sub(r'[\s\u200b\u200c\u200d\ufeff]+', '', text)
        thai_to_ar = str.maketrans('๑๒๓๔๕๖๗๘๙๐', '1234567890')
        return text.translate(thai_to_ar)

    def _load_db(self):
        csv_path = self.db_path.replace('.xlsx', '.csv')
        if not os.path.exists(self.db_path) and os.path.exists(csv_path):
            try: pd.read_csv(csv_path).to_excel(self.db_path, index=False)
            except: pass

        if os.path.exists(self.db_path):
            try:
                self.db = pd.read_excel(self.db_path)
                for col in ['item', 'v_name', 'code']:
                    if col not in self.db.columns: self.db[col] = ''
                
                self.db['item'] = self.db['item'].fillna('').astype(str).str.strip().replace('nan', '')
                self.db['v_name'] = self.db['v_name'].fillna('').astype(str).str.strip().replace('nan', '')
                self.db['code'] = self.db['code'].fillna('').astype(str).str.strip().replace('nan', '')
                
                self.db['item_norm'] = self.db['item'].apply(self._normalize_for_match)
                self.db['v_name_norm'] = self.db['v_name'].apply(self._normalize_for_match)

                self.db['_mk'] = self.db['item_norm'] + self.db['v_name_norm']
                self.db = self.db.drop_duplicates(subset=['_mk'], keep='last').drop(columns=['_mk'])

            except Exception as e:
                print(f"Error loading DB: {e}")
                self.db = pd.DataFrame(columns=['item', 'v_name', 'code', 'item_norm', 'v_name_norm'])
        else:
            self.db = pd.DataFrame(columns=['item', 'v_name', 'code', 'item_norm', 'v_name_norm'])

    def clean_thai_text(self, text):
        if not text: return ""
        text = str(text)
        text = unicodedata.normalize('NFKC', text)
        text = thai_normalize(text)
        noise = ["Order ID", "Total Amount", "No.", "Seller SKU", "Shopee Order No", "Package ID", "COD", "PICK-UP"]
        for n in noise:
            text = re.sub(rf'{n}.*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'([\u0e00-\u0e7f])\s+([\u0e00-\u0e7f])', r'\1\2', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def add_labels_to_pdf(self, input_pdf, output_pdf, orders_to_label):
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        font_size = 18
        line_height = font_size * 1.2
        
        is_iship = "iship" in os.path.basename(input_pdf).lower()
        
        for i, page in enumerate(reader.pages):
            page_num = i + 1
            if page_num in orders_to_label:
                packet = io.BytesIO()
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
                can = canvas.Canvas(packet, pagesize=(w, h))
                
                if self.font_path:
                    try:
                        pdfmetrics.registerFont(TTFont('ThaiFont', self.font_path))
                        can.setFont('ThaiFont', font_size)
                    except: can.setFont('Helvetica-Bold', font_size)
                else: can.setFont('Helvetica-Bold', font_size)
                
                for order in orders_to_label[page_num]:
                    raw_code = order['code'].replace("\\n", "\n").strip()
                    qty = order.get('qty', 1)
                    
                    lines = []
                    for part in raw_code.split("\n"):
                        part = part.strip()
                        if part:
                            part_with_qty = f"{part}......x{qty}"
                            lines.extend(textwrap.wrap(part_with_qty, width=100))
                    
                    max_line_w = max([can.stringWidth(line, can._fontname, font_size) for line in lines])
                    box_w = max_line_w + 20
                    box_h = (len(lines) * line_height) + 10
                    
                    if is_iship or "ship.pdf" in input_pdf.lower():
                        rect_x = w - box_w - 15; rect_y = 15
                    else:
                        rect_y = (h / 2) + 15 if order['y_pos'] < h / 2 else 15
                        rect_x = w - box_w - 15
                    
                    can.setFillColorRGB(1, 1, 0)
                    can.setStrokeColorRGB(1, 0, 0)
                    can.setLineWidth(1.2)
                    can.rect(rect_x, rect_y, box_w, box_h, fill=1, stroke=1)
                    can.setFillColorRGB(1, 0, 0)
                    for j, line in enumerate(reversed(lines)):
                        can.drawString(rect_x + 5, rect_y + 8 + (j * line_height), line)
                
                can.save()
                packet.seek(0)
                overlay_reader = PdfReader(packet)
                page.merge_page(overlay_reader.pages[0])
                writer.add_page(page)
            elif not is_iship:
                writer.add_page(page)
                
        with open(output_pdf, "wb") as f:
            writer.write(f)
        return output_pdf

    def extract_order_data(self, pdf_path):
        self._load_db() 
        all_results = []
        fname_low = os.path.basename(pdf_path).lower()
        try:
            doc = fitz.open(pdf_path)
            
            if "iship" in fname_low:
                pending_label = None
                for p_num in range(len(doc)):
                    page = doc[p_num]
                    full_text = page.get_text("text").strip()
                    match = re.search(r'(?:หมายเหตุ|Remark|หมายเหต|หมายเหด)[:\s]*(.*)', full_text, re.IGNORECASE)
                    v_name = match.group(1).strip() if match else ""
                    v_name = re.sub(r'[^\w\s\+\-ก-๙]', '', v_name).strip()
                    
                    is_label = len(full_text) > 150 or any(x in full_text for x in ["ชื่อผู้", "รายการ", "TH", "Tracking"])
                    if is_label:
                        res = {
                            'page': p_num + 1, 'file': os.path.basename(pdf_path),
                            'item': 'iShip Label', 'v_name': v_name, 'qty': 1, 
                            'code': f"MANUAL {v_name}" if v_name else "MANUAL", 
                            'y_pos': 0, 'file_path': pdf_path
                        }
                        all_results.append(res)
                        pending_label = res
                    elif v_name and pending_label:
                        pending_label['v_name'] = v_name
                        pending_label['code'] = f"MANUAL {v_name}"
                doc.close()
                if all_results: return all_results

            for page_num, page in enumerate(doc, 1):
                dict_text = page.get_text("dict")
                spans = []
                for block in dict_text["blocks"]:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                spans.append({
                                    "text": span["text"], "bbox": span["bbox"],
                                    "y_center": (span["bbox"][1] + span["bbox"][3]) / 2,
                                    "x_center": (span["bbox"][0] + span["bbox"][2]) / 2
                                })
                if not spans: continue
                headers = []
                spans_sorted = sorted(spans, key=lambda x: x["y_center"])
                for i, s in enumerate(spans_sorted):
                    txt = s["text"].lower().strip()
                    if s["y_center"] > 500 and len(headers) == 0: continue 
                    if "item" == txt:
                        h_info = {"item_s": s, "qty_s": None, "v_name_s": None}
                        for s2 in spans:
                            if s2 == s: continue
                            txt2 = s2["text"].lower().strip()
                            if abs(s2["y_center"] - s["y_center"]) < 20: 
                                if "qty" == txt2 or "จำนวน" == txt2: h_info["qty_s"] = s2
                                if any(x == txt2 for x in ["v name", "v_name", "variant", "รุ่น"]): h_info["v_name_s"] = s2
                        if h_info["qty_s"]: headers.append(h_info)
                
                headers.sort(key=lambda x: x["item_s"]["y_center"])
                for i, h_group in enumerate(headers):
                    header_s = h_group["item_s"]; header_y = header_s["y_center"]
                    item_start_x = header_s["bbox"][0]
                    qty_start_x = h_group["qty_s"]["bbox"][0]
                    v_start_x = h_group["v_name_s"]["bbox"][0] if h_group["v_name_s"] else (qty_start_x + 30)
                    v_on_left = v_start_x < qty_start_x
                    next_header_y = headers[i+1]["item_s"]["y_center"] if i+1 < len(headers) else page.rect.height
                    label_spans = [s for s in spans if header_y + 2 < s["y_center"] < next_header_y - 2]
                    item_parts, v_parts, qty_parts = [], [], []
                    for s in label_spans:
                        t = s["text"].strip()
                        if not t or "Order ID" in t or "Seller SKU" in t: continue
                        x_start = s["bbox"][0]
                        if v_on_left:
                            if x_start > (v_start_x - 10) and x_start < (qty_start_x - 10): v_parts.append(t)
                            elif x_start > (qty_start_x - 10):
                                if t.isdigit() or (t.startswith('x') and t[1:].isdigit()): qty_parts.append(t)
                            else: item_parts.append(t)
                        else:
                            if x_start > (qty_start_x - 10) and x_start < (v_start_x - 10):
                                if t.isdigit() or (t.startswith('x') and t[1:].isdigit()): qty_parts.append(t)
                            elif x_start > (v_start_x - 10): v_parts.append(t)
                            else: item_parts.append(t)
                    
                    item_ext = self.clean_thai_text(" ".join(item_parts))
                    v_ext = self.clean_thai_text(" ".join(v_parts))
                    qty = 1
                    try:
                        qty_str = "".join(qty_parts).replace('x', '')
                        if qty_str.isdigit(): qty = int(qty_str)
                    except: pass
                    
                    if item_ext:
                        item_norm = self._normalize_for_match(item_ext)
                        v_norm = self._normalize_for_match(v_ext)
                        
                        best_code = "NOT FOUND"
                        best_score = -9999
                        
                        for _, db_row in self.db.iterrows():
                            db_i_norm = db_row.get('item_norm', '')
                            db_v_norm = db_row.get('v_name_norm', '')
                            if not db_i_norm: continue 

                            item_score = 0
                            if db_i_norm == item_norm: item_score = 1000 
                            elif db_i_norm in item_norm or item_norm in db_i_norm: item_score = len(db_i_norm) 
                            else: continue 
                                
                            v_score = 0
                            if not v_ext and not db_v_norm: v_score = 1000 
                            elif db_v_norm == v_norm: v_score = 1000 
                            elif not db_v_norm: v_score = 0 
                            elif db_v_norm in v_norm or v_norm in db_v_norm: v_score = len(db_v_norm) 
                            else: continue 
                                
                            matched_code = str(db_row.get('code', '')).strip()
                            total_score = item_score + v_score
                            
                            if not matched_code or matched_code.upper() == "NOT FOUND":
                                total_score -= 5000 
                            
                            if total_score > best_score:
                                best_score = total_score
                                best_code = matched_code
                            elif total_score == best_score:
                                if matched_code.upper() != "NOT FOUND":
                                    best_code = matched_code
                                
                        code = best_code if best_code else "NOT FOUND"
                                    
                        all_results.append({
                            'page': page_num, 'file': os.path.basename(pdf_path),
                            'item': item_ext, 'v_name': v_ext, 'qty': qty,
                            'code': code, 'y_pos': header_y, 'file_path': pdf_path
                        })

            # ปิดไฟล์ PDF แล้วส่งค่ากลับเลย (หั่นระบบ Tesseract ออกไปแล้ว)
            doc.close()
            return all_results
        except Exception as e: print(f"Extraction error: {e}"); return []

    def import_excel_db(self, excel_path):
        try:
            df = pd.read_excel(excel_path); new_data = []
            for _, row in df.iterrows():
                item_val, v_val, code_val = "", "", ""
                for col in df.columns:
                    c_low = str(col).lower()
                    if 'ชื่อ sku' in c_low or 'สินค้า' in c_low or 'item' in c_low: item_val = str(row[col])
                    if 'เลข sku' in c_low or 'variant' in c_low or 'รุ่น' in c_low: v_val = str(row[col])
                    if 'code' in c_low or 'รหัส' in c_low: code_val = str(row[col])
                if item_val and v_val:
                    new_data.append({'item': self.clean_thai_text(item_val), 'v_name': self.clean_thai_text(v_val), 'code': str(code_val if code_val else v_val).strip()})
            if new_data:
                import_df = pd.DataFrame(new_data)
                if os.path.exists(self.db_path):
                    existing = pd.read_excel(self.db_path)
                    updated = pd.concat([existing, import_df], ignore_index=True)
                    updated['_mk'] = updated['item'].apply(self._normalize_for_match) + updated['v_name'].apply(self._normalize_for_match)
                    updated = updated.drop_duplicates(subset=['_mk'], keep='last').drop(columns=['_mk'])
                    updated.to_excel(self.db_path, index=False)
                else: import_df.to_excel(self.db_path, index=False)
                self._load_db(); return True, len(new_data)
            return False, 0
        except Exception as e: return False, str(e)

    def save_to_db(self, item, v_name, code, old_item=None, old_v_name=None):
        if not item: return False, "No item name"
        item = self.clean_thai_text(item)
        v_name = self.clean_thai_text(v_name)
        code = str(code).strip() 
        
        new_row = {'item': item, 'v_name': v_name, 'code': code}
        
        try:
            if os.path.exists(self.db_path):
                df = pd.read_excel(self.db_path)
                for col in ['item', 'v_name', 'code']:
                    if col not in df.columns: df[col] = ''
                df['item'] = df['item'].fillna('').astype(str)
                df['v_name'] = df['v_name'].fillna('').astype(str)
                df['code'] = df['code'].fillna('').astype(str)
                
                t_item = old_item if old_item is not None else item
                t_v = old_v_name if old_v_name is not None else v_name
                
                t_i_norm = self._normalize_for_match(t_item)
                t_v_norm = self._normalize_for_match(t_v)
                new_i_norm = self._normalize_for_match(item)
                new_v_norm = self._normalize_for_match(v_name)

                df['_match_i'] = df['item'].apply(self._normalize_for_match)
                df['_match_v'] = df['v_name'].apply(self._normalize_for_match)

                mask1 = (df['_match_i'] == t_i_norm) & (df['_match_v'] == t_v_norm)
                mask2 = (df['_match_i'] == new_i_norm) & (df['_match_v'] == new_v_norm)
                mask = mask1 | mask2
                
                df = df[~mask].copy() 
                df = df.drop(columns=['_match_i', '_match_v'], errors='ignore')
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            else:
                df = pd.DataFrame([new_row])
            
            df['_match_i'] = df['item'].apply(self._normalize_for_match)
            df['_match_v'] = df['v_name'].apply(self._normalize_for_match)
            df = df.drop_duplicates(subset=['_match_i', '_match_v'], keep='last')
            df = df.drop(columns=['_match_i', '_match_v'], errors='ignore')

            cols = ['item', 'v_name', 'code']
            other_cols = [c for c in df.columns if c not in cols and not c.startswith('_')]
            df = df[cols + other_cols]

            df.to_excel(self.db_path, index=False)
            self._load_db()
            return True, ""
        except Exception as e:
            return False, f"กรุณาปิดไฟล์ Excel ก่อนเซฟ: {e}"

    def delete_from_db(self, item, v_name):
        if os.path.exists(self.db_path):
            try:
                df = pd.read_excel(self.db_path)
                for col in ['item', 'v_name']:
                    if col not in df.columns: df[col] = ''
                df['item'] = df['item'].fillna('').astype(str)
                df['v_name'] = df['v_name'].fillna('').astype(str)

                t_i_norm = self._normalize_for_match(item)
                t_v_norm = self._normalize_for_match(v_name)
                
                df['_match_i'] = df['item'].apply(self._normalize_for_match)
                df['_match_v'] = df['v_name'].apply(self._normalize_for_match)

                mask = (df['_match_i'] == t_i_norm) & (df['_match_v'] == t_v_norm)
                updated_db = df[~mask].copy()
                
                updated_db = updated_db.drop(columns=['_match_i', '_match_v'], errors='ignore')
                updated_db.to_excel(self.db_path, index=False)
                self._load_db(); return True
            except: return False
        return False

    def clear_db(self):
        try:
            pd.DataFrame(columns=['item', 'v_name', 'code']).to_excel(self.db_path, index=False)
            self._load_db(); return True
        except: return False