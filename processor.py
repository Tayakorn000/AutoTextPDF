import fitz # PyMuPDF
import pandas as pd
import os
import re
from collections import defaultdict
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
        is_iship = "iship" in os.path.basename(input_pdf).lower()
        
        for i, page in enumerate(reader.pages):
            page_num = i + 1
            if page_num in orders_to_label:
                packet = io.BytesIO()
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
                can = canvas.Canvas(packet, pagesize=(w, h))
                
                # Group items by zone (Top/Bottom half) so items in the same order share a box
                zones = defaultdict(list)
                for order in orders_to_label[page_num]:
                    if is_iship or "ship.pdf" in os.path.basename(input_pdf).lower():
                        zone_key = "full"
                    else:
                        zone_key = "top" if order['y_pos'] < h/2 else "bottom"
                    zones[zone_key].append(order)
                
                for zone_key, zone_items in zones.items():
                    all_lines_data = []
                    zone_items.sort(key=lambda x: x['y_pos'])
                    
                    for idx, itm in enumerate(zone_items):
                        raw_code = itm['code'].replace("\\n", "\n").strip()
                        qty = itm.get('qty', 1)
                        if idx > 0:
                            all_lines_data.append({'text': "", 'qty': 0, 'is_spacer': True})
                        
                        # Append qty to EVERY line of this item
                        item_lines = [l.strip() for l in raw_code.split("\n") if l.strip()]
                        if not item_lines: continue
                        for line_text in item_lines:
                            final_text = f"{line_text}......x{qty}"
                            all_lines_data.append({
                                'text': final_text, 'qty': qty, 'is_spacer': False
                            })

                    if not all_lines_data: continue
                    font_size = 18; max_box_w = 230
                    while font_size > 7:
                        if self.font_path:
                            try:
                                pdfmetrics.registerFont(TTFont('ThaiFont', self.font_path))
                                can.setFont('ThaiFont', font_size)
                            except: can.setFont('Helvetica-Bold', font_size)
                        else: can.setFont('Helvetica-Bold', font_size)
                        max_w = max([can.stringWidth(d['text'], can._fontname, font_size) for d in all_lines_data if not d['is_spacer']])
                        if max_w < max_box_w: break
                        font_size -= 1
                    
                    line_h = font_size * 1.2; box_w = max_w + 20; box_h = 10
                    for d in all_lines_data: box_h += (line_h * 0.5 if d['is_spacer'] else line_h)
                    
                    if zone_key == "full": rx = w - box_w - 15; ry = 15
                    elif zone_key == "top": rx = w - box_w - 15; ry = (h / 2) + 15
                    else: rx = w - box_w - 15; ry = 15
                    
                    can.setFillColorRGB(1, 1, 0); can.setStrokeColorRGB(1, 0, 0); can.setLineWidth(1.2)
                    can.rect(rx, ry, box_w, box_h, fill=1, stroke=1)
                    can.setFillColorRGB(1, 0, 0); curr_y = ry + 8
                    for d in reversed(all_lines_data):
                        if d['is_spacer']: curr_y += line_h * 0.5; continue
                        can.drawString(rx + 5, curr_y, d['text'])
                        if d['qty'] > 1:
                            can.setStrokeColorRGB(1, 0, 0); can.setLineWidth(1.5)
                            qw = can.stringWidth(f"x{d['qty']}", can._fontname, font_size)
                            tw = can.stringWidth(d['text'], can._fontname, font_size)
                            cx = rx + 5 + tw - (qw / 2); cy = curr_y + (font_size * 0.35)
                            can.circle(cx, cy, font_size * 0.7, fill=0, stroke=1)
                        curr_y += line_h
                
                can.save()
                packet.seek(0)
                overlay_reader = PdfReader(packet)
                page.merge_page(overlay_reader.pages[0])
                writer.add_page(page)
            else:
                writer.add_page(page)
                
        with open(output_pdf, "wb") as f:
            writer.write(f)
        return output_pdf

    def extract_order_data(self, pdf_path):
        self._load_db() 
        all_results = []
        fname_low = os.path.basename(pdf_path).lower()
        active_headers = None # Track headers across pages for overflows
        
        try:
            doc = fitz.open(pdf_path)
            
            if "iship" in fname_low:
                # ... (iship logic remains the same)
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
                    # Look for Item header
                    if "item" == txt or "รายการ" == txt:
                        h_info = {"item_s": s, "qty_s": None, "v_name_s": None}
                        for s2 in spans:
                            if s2 == s: continue
                            txt2 = s2["text"].lower().strip()
                            if abs(s2["y_center"] - s["y_center"]) < 25: 
                                if "qty" == txt2 or "จำนวน" == txt2: h_info["qty_s"] = s2
                                if any(x == txt2 for x in ["v name", "v_name", "variant", "รุ่น", "แบบ"]): h_info["v_name_s"] = s2
                        if h_info["qty_s"]: headers.append(h_info)
                
                # If no headers found but we had them on previous page, create a virtual header for continuation
                if not headers and active_headers:
                    # Only if there's text at the top of the page (likely a continuation)
                    if any(s["y_center"] < 200 for s in spans):
                        headers = [{
                            "item_s": {"bbox": active_headers["item_s"]["bbox"], "y_center": 0},
                            "qty_s": active_headers["qty_s"],
                            "v_name_s": active_headers["v_name_s"]
                        }]

                # Filter and deduplicate headers
                unique_headers = []
                for h in sorted(headers, key=lambda x: x["item_s"]["y_center"]):
                    if not unique_headers or abs(h["item_s"]["y_center"] - unique_headers[-1]["item_s"]["y_center"]) > 20:
                        unique_headers.append(h)
                
                if unique_headers:
                    active_headers = unique_headers[-1] # Save the last header for the next page overflow

                for i, h_group in enumerate(unique_headers):
                    header_s = h_group["item_s"]; header_y = header_s["y_center"]
                    qty_x = h_group["qty_s"]["bbox"][0]
                    v_x = h_group["v_name_s"]["bbox"][0] if h_group["v_name_s"] else (qty_x + 40)
                    v_on_left = v_x < qty_x
                    
                    next_h_y = unique_headers[i+1]["item_s"]["y_center"] if i+1 < len(unique_headers) else page.rect.height
                    
                    # Zone spans
                    zone_spans = sorted([s for s in spans if header_y + 5 < s["y_center"] < next_h_y - 5], key=lambda x: x["y_center"])
                    
                    rows = []
                    if zone_spans:
                        cur_row = [zone_spans[0]]
                        for s_idx in range(1, len(zone_spans)):
                            if abs(zone_spans[s_idx]["y_center"] - zone_spans[s_idx-1]["y_center"]) < 6:
                                cur_row.append(zone_spans[s_idx])
                            else:
                                rows.append(cur_row)
                                cur_row = [zone_spans[s_idx]]
                        rows.append(cur_row)
                    
                    items_found = []
                    cur_item = {'item': [], 'v': [], 'qty': None, 'y': None}
                    
                    for row in rows:
                        r_item, r_v, r_qty = [], [], None
                        ry = row[0]["y_center"]
                        row = sorted(row, key=lambda x: x["bbox"][0])
                        
                        for s in row:
                            t = s["text"].strip()
                            # Filter out common PDF noise
                            if not t or any(x in t for x in ["Order ID", "Seller SKU", "Package ID", "Total Amount", "Page ", "Date"]): continue
                            x = s["bbox"][0]
                            
                            # Check if this span is in the Qty column
                            is_in_qty_col = False
                            if v_on_left:
                                if x > (qty_x - 10): is_in_qty_col = True
                            else:
                                if (qty_x - 10) < x < (v_x - 5): is_in_qty_col = True
                            
                            if is_in_qty_col:
                                # Look for quantity patterns: "2", "x2", "x 2"
                                m = re.search(r'(?:x\s*)?(\d+)', t, re.IGNORECASE)
                                if m and (len(t) < 5 or t.lower().startswith('x')): # Ensure it's not a long string with a number
                                    r_qty = int(m.group(1))
                                else:
                                    if v_on_left: r_v.append(t)
                                    else: r_item.append(t)
                            elif not v_on_left and x > (v_x - 10):
                                r_v.append(t)
                            elif v_on_left and (v_x - 10) < x < (qty_x - 5):
                                r_v.append(t)
                            else:
                                r_item.append(t)
                        
                        # LOGIC: Only split into a new item if we see a NEW qty and the current one ALREADY has a qty
                        if r_qty is not None:
                            if cur_item['qty'] is not None:
                                # Current item is complete, save it
                                items_found.append(cur_item)
                                cur_item = {'item': r_item, 'v': r_v, 'qty': r_qty, 'y': ry}
                            else:
                                # Current item was waiting for a qty
                                cur_item['item'].extend(r_item)
                                cur_item['v'].extend(r_v)
                                cur_item['qty'] = r_qty
                                if cur_item['y'] is None: cur_item['y'] = ry
                        else:
                            # No qty on this line, just more text for whatever we are building
                            cur_item['item'].extend(r_item)
                            cur_item['v'].extend(r_v)
                            if cur_item['y'] is None: cur_item['y'] = ry
                            
                    # Final item cleanup
                    if cur_item['item'] or cur_item['v']:
                        items_found.append(cur_item)

                    for itm in items_found:
                        # VALIDATION: An item MUST have a quantity to be valid
                        if itm['qty'] is None: continue
                        
                        item_t = self.clean_thai_text(" ".join(itm['item']))
                        v_t = self.clean_thai_text(" ".join(itm['v']))
                        
                        if not item_t and not v_t: continue
                        if len(item_t) < 2 and not v_t: continue 
                        
                        qty = itm['qty']
                        y_pos = itm['y']
                        
                        item_n, v_n = self._normalize_for_match(item_t), self._normalize_for_match(v_t)
                        best_c, best_s = "NOT FOUND", -9999
                        
                        for _, db_row in self.db.iterrows():
                            db_i, db_v = db_row.get('item_norm', ''), db_row.get('v_name_norm', '')
                            if not db_i: continue 
                            i_s = 1000 if db_i == item_n else (len(db_i) if db_i in item_n or item_n in db_i else -1)
                            if i_s < 0: continue
                            v_s = 1000 if (not v_t and not db_v) or db_v == v_n else (len(db_v) if db_v and (db_v in v_n or v_n in db_v) else 0)
                            score = i_s + v_s
                            code = str(db_row.get('code', '')).strip()
                            if not code or code.upper() == "NOT FOUND": score -= 5000 
                            if score > best_s: best_s, best_c = score, code
                                    
                        all_results.append({
                            'page': page_num, 'file': os.path.basename(pdf_path),
                            'item': item_t, 'v_name': v_t, 'qty': qty,
                            'code': best_c, 'y_pos': y_pos, 'file_path': pdf_path
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