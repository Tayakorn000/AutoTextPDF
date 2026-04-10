import os
import sys
import datetime
from collections import defaultdict
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
                               QTextEdit, QSplitter, QFileDialog, QMessageBox, QDialog, QInputDialog,
                               QAbstractItemView, QLineEdit, QFormLayout, QComboBox)
from PySide6.QtCore import Qt, QTimer, Slot
from PySide6.QtGui import QKeySequence, QShortcut, QClipboard, QAction

from processor import PDFProcessor

class MultilineInputDialog(QDialog):
    def __init__(self, parent=None, title="Edit Content", initial_value="", show_toolbar=True):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(600, 450)
        
        # Apply dark styling to dialog
        self.setStyleSheet("""
            QDialog { background-color: #242424; color: white; }
            QLabel { color: white; font-weight: bold; font-size: 13px; }
            QTextEdit { background-color: #1a1a1a; color: white; border: 1px solid #555; font-size: 14px; font-family: 'Tahoma'; }
            QPushButton { background-color: #343a40; color: white; border: 1px solid #555; padding: 6px; border-radius: 4px; }
            QPushButton:hover { background-color: #495057; }
            QComboBox { background-color: #1a1a1a; color: white; border: 1px solid #555; padding: 3px; }
        """)
        
        layout = QVBoxLayout(self)
        
        # --- Toolbar (Conditional) ---
        if show_toolbar:
            toolbar = QHBoxLayout()
            
            btn_b = QPushButton("B")
            btn_b.setFixedWidth(35); btn_b.setStyleSheet("font-weight: bold; font-size: 16px;")
            btn_b.clicked.connect(lambda: self._wrap_selection("<b>", "</b>"))
            
            btn_i = QPushButton("I")
            btn_i.setFixedWidth(35); btn_i.setStyleSheet("font-style: italic; font-size: 16px;")
            btn_i.clicked.connect(lambda: self._wrap_selection("<i>", "</i>"))
            
            btn_u = QPushButton("U")
            btn_u.setFixedWidth(35); btn_u.setStyleSheet("text-decoration: underline; font-size: 16px;")
            btn_u.clicked.connect(lambda: self._wrap_selection("<u>", "</u>"))
            
            toolbar.addWidget(btn_b)
            toolbar.addWidget(btn_i)
            toolbar.addWidget(btn_u)
            toolbar.addSpacing(15)
            
            toolbar.addWidget(QLabel("Font Size:"))
            self.size_box = QComboBox()
            self.size_box.addItems(["16", "20", "24", "30", "36", "40", "50", "60", "72"])
            self.size_box.setCurrentText("20")
            self.size_box.setFixedWidth(65)

            toolbar.addWidget(self.size_box)
            
            btn_set_size = QPushButton("Apply Size")
            btn_set_size.clicked.connect(self._apply_size)
            toolbar.addWidget(btn_set_size)
            
            toolbar.addStretch()
            layout.addLayout(toolbar)
            
            help_label = QLabel("Tip: ลากคลุมข้อความแล้วกดปุ่มด้านบนเพื่อจัดรูปแบบ")
            help_label.setStyleSheet("color: #888; font-size: 11px; font-weight: normal;")
            layout.addWidget(help_label)
        
        # --- Editor ---
        self.textbox = QTextEdit()
        self.textbox.setPlainText(initial_value)
        layout.addWidget(self.textbox)
        
        # --- Actions ---
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("Save & Update DB")
        self.btn_save.setStyleSheet("background-color: #28a745; font-weight: bold; min-height: 35px;")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setMinimumHeight(35)
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)
        
        self.btn_save.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        
        self.result = None

    def _wrap_selection(self, start, end):
        cursor = self.textbox.textCursor()
        if cursor.hasSelection():
            txt = cursor.selectedText()
            cursor.insertText(f"{start}{txt}{end}")
        else:
            cursor.insertText(f"{start}{end}")
            for _ in range(len(end)): cursor.movePosition(cursor.Left)
            self.textbox.setTextCursor(cursor)

    def _apply_size(self):
        size = self.size_box.currentText()
        self._wrap_selection(f'<font size="{size}">', '</font>')

    def accept(self):
        self.result = self.textbox.toPlainText()
        super().accept()

class DropZone(QLabel):
    def __init__(self, main_window):
        super().__init__("Drag PDFs here or Click to Select")
        self.main_window = main_window
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #3b8ed0;
                border-radius: 10px;
                background-color: #1a1a1a;
                font-size: 16px;
                font-weight: bold;
                color: #ffffff;
            }
        """)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
            
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().lower().endswith('.pdf')]
        for f in files:
            self.main_window._process_pdf(f)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window._select_file()

class PDFLabelerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Product Labeler (PySide6 Edition)")
        self.resize(1450, 850)
        self.processor = PDFProcessor()
        self.file_to_order_meta = {}
        self._is_populating = False
        # ...
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #242424; color: #ffffff; }
            QPushButton { background-color: #343a40; color: white; border: 1px solid #555; padding: 5px; border-radius: 5px; }
            QPushButton:hover { background-color: #495057; }
            QTableWidget { background-color: #2b2b2b; alternate-background-color: #3d3d3d; color: white; border: none; font-size: 13px; }
            QHeaderView::section { background-color: #3d3d3d; color: white; padding: 4px; font-weight: bold; border: 1px solid #222; }
            QTextEdit { background-color: #1a1a1a; color: white; border: 1px solid #555; }
        """)
        
        self._create_widgets()

    def _create_widgets(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Splitter for layout
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # --- Left Sidebar ---
        sidebar = QWidget()
        sidebar.setMinimumWidth(260)
        sidebar.setMaximumWidth(300)
        sidebar_layout = QVBoxLayout(sidebar)
        
        title_lbl = QLabel("PDF Product Labeler")
        font = title_lbl.font()
        font.setPointSize(16)
        font.setBold(True)
        title_lbl.setFont(font)
        title_lbl.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(title_lbl)
        
        sidebar_layout.addSpacing(20)
        
        btn_open_db = QPushButton("Open DB (Excel)")
        btn_open_db.setMinimumHeight(40)
        btn_open_db.clicked.connect(self._open_db_file)
        sidebar_layout.addWidget(btn_open_db)
        
        btn_reload_db = QPushButton("Reload DB")
        btn_reload_db.setMinimumHeight(40)
        btn_reload_db.setStyleSheet("background-color: #17a2b8; font-weight: bold;")
        btn_reload_db.clicked.connect(self._reload_db)
        sidebar_layout.addWidget(btn_reload_db)
        
        sidebar_layout.addSpacing(10)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        font = self.log_text.font()
        font.setPointSize(10)
        self.log_text.setFont(font)
        sidebar_layout.addWidget(self.log_text)
        
        # --- Main Content ---
        main_content = QWidget()
        mc_layout = QVBoxLayout(main_content)
        
        # Drop Zone
        self.drop_zone = DropZone(self)
        self.drop_zone.setMinimumHeight(150)
        self.drop_zone.setMaximumHeight(200)
        mc_layout.addWidget(self.drop_zone)
        
        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(["ID", "File", "Page", "Item", "Variant", "Qty", "Code", "Manual", "Manual Text"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        
        # *** ปรับความกว้างคอลัมน์ใหม่ ***
        header = self.table.horizontalHeader()
        header.resizeSection(0, 30)   # ID
        header.resizeSection(1, 80)   # File
        header.resizeSection(2, 40)   # Page
        header.resizeSection(3, 300)  # Item
        header.resizeSection(4, 200)  # Variant
        header.resizeSection(5, 45)   # Qty
        header.resizeSection(6, 150)  # Code
        header.resizeSection(7, 60)   # Manual
        header.setSectionResizeMode(8, QHeaderView.Stretch) # Manual Text
        
        self.table.verticalHeader().setVisible(False)
        self.table.cellDoubleClicked.connect(self._on_double_click)
        self.table.itemChanged.connect(self._on_item_changed)
        mc_layout.addWidget(self.table)
        
        # Action Buttons
        btn_layout = QHBoxLayout()
        
        btn_clear = QPushButton("Clear List")
        btn_clear.setMinimumHeight(35)
        btn_clear.setStyleSheet("background-color: #dc3545; font-weight: bold;")
        btn_clear.clicked.connect(self._clear_table)
        btn_layout.addWidget(btn_clear)
        
        btn_update = QPushButton("Update DB (Manual)")
        btn_update.setMinimumHeight(35)
        btn_update.setStyleSheet("background-color: #28a745; font-weight: bold;")
        btn_update.clicked.connect(self._save_selected_to_db)
        btn_layout.addWidget(btn_update)
        
        btn_layout.addStretch()
        
        btn_print = QPushButton("PRINT ALL & CLEAR (Shift+B)")
        btn_print.setMinimumHeight(45)
        btn_print.setMinimumWidth(280)
        font = btn_print.font()
        font.setPointSize(12)
        font.setBold(True)
        btn_print.setFont(font)
        btn_print.setStyleSheet("background-color: #007bff;")
        btn_print.clicked.connect(self._label_pdfs)
        btn_layout.addWidget(btn_print)
        
        mc_layout.addLayout(btn_layout)
        
        splitter.addWidget(sidebar)
        splitter.addWidget(main_content)
        splitter.setSizes([260, 1190])
        
        # Shortcuts
        QShortcut(QKeySequence("Shift+B"), self, self._label_pdfs)
        
        # Universal Copy/Paste actions
        copy_act = QAction("Copy", self)
        copy_act.setShortcut(QKeySequence.Copy)
        copy_act.triggered.connect(self._on_copy)
        self.table.addAction(copy_act)
        
        paste_act = QAction("Paste", self)
        paste_act.setShortcut(QKeySequence.Paste)
        paste_act.triggered.connect(self._on_paste)
        self.table.addAction(paste_act)
        
        self.table.setContextMenuPolicy(Qt.ActionsContextMenu)

    def _log(self, message):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {message}")

    def _on_copy(self):
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges: return
        
        rows = sorted(list(set([r.topRow() for r in selected_ranges])))
        if rows:
            row = rows[0]
            item_val = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            var_val = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            qty_val = self.table.item(row, 5).text() if self.table.item(row, 5) else ""
            code_val = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            copy_str = f"{item_val}\t{var_val}\t{qty_val}\t{code_val}"
            QApplication.clipboard().setText(copy_str)
            self._log("Copied data to clipboard.")

    def _on_paste(self):
        cb = QApplication.clipboard().text()
        if not cb: return
        
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges: return
        rows = sorted(list(set([r.topRow() for r in selected_ranges])))
        
        for row in rows:
            old_item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            old_v = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            
            if "\t" in cb:
                parts = cb.split("\t")
                if len(parts) >= 4:
                    self.table.setItem(row, 3, QTableWidgetItem(parts[0]))
                    self.table.setItem(row, 4, QTableWidgetItem(parts[1]))
                    self.table.setItem(row, 5, QTableWidgetItem(parts[2]))
                    self.table.setItem(row, 6, QTableWidgetItem(parts[3]))
            else:
                self.table.setItem(row, 6, QTableWidgetItem(cb))
                
            n_item = self.table.item(row, 3).text()
            n_v = self.table.item(row, 4).text()
            n_code = self.table.item(row, 6).text()
            
            success, msg = self.processor.save_to_db(n_item, n_v, n_code, old_item=old_item, old_v_name=old_v)
            if not success:
                QMessageBox.critical(self, "Save Error", f"Failed to save to database.\nError: {msg}")
                return
        
        db_fname = os.path.basename(self.processor.db_path)
        self._log(f"Auto-Saved mapping to {db_fname} (Fuzzy Match).")

    def _reload_db(self):
        self.processor._load_db()
        self._log("Reloaded database from Excel.")
        QMessageBox.information(self, "DB Reloaded", "โหลดข้อมูลจาก Excel เรียบร้อยแล้ว")

    def _select_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDF Orders", "", "PDF files (*.pdf)")
        if files:
            for f in files: self._process_pdf(f)

    def _process_pdf(self, file_path):
        fname = os.path.basename(file_path)
        self._log(f"Mapping {fname}...")
        results = self.processor.extract_order_data(file_path)
        if not results:
            self._log(f"No orders in {fname}")
            return
            
        self._is_populating = True
        self.table.setSortingEnabled(False)
        for res in results:
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            rid = str(row + 1)
            
            self.table.setItem(row, 0, QTableWidgetItem(rid))
            self.table.setItem(row, 1, QTableWidgetItem(fname))
            self.table.setItem(row, 2, QTableWidgetItem(str(res['page'])))
            self.table.setItem(row, 3, QTableWidgetItem(res['item']))
            self.table.setItem(row, 4, QTableWidgetItem(res['v_name']))
            self.table.setItem(row, 5, QTableWidgetItem(str(res['qty'])))
            self.table.setItem(row, 6, QTableWidgetItem(res['code']))
            
            # Manual Checkbox
            manual_item = QTableWidgetItem()
            manual_item.setCheckState(Qt.Checked if res.get('has_manual') else Qt.Unchecked)
            self.table.setItem(row, 7, manual_item)
            
            self.table.setItem(row, 8, QTableWidgetItem(res.get('manual_text', '')))
            
            # Center alignment for ID, Page, Qty, Manual
            for col in (0, 2, 5, 7):
                item = self.table.item(row, col)
                if item: item.setTextAlignment(Qt.AlignCenter)
            
            self.file_to_order_meta[row] = {
                'path': file_path, 
                'page': res['page'],
                'y_pos': res['y_pos'],
                'qty': res['qty'],
                'has_manual': res.get('has_manual', 0),
                'manual_text': res.get('manual_text', '')
            }
        self._is_populating = False

    def _open_db_file(self):
        try:
            if os.name == 'nt': os.startfile(self.processor.db_path)
            else: os.system(f"open '{self.processor.db_path}'")
        except Exception as e:
            self._log(f"Cannot open DB: {e}")

    def _clear_table(self):
        self._is_populating = True
        self.table.setRowCount(0)
        self.file_to_order_meta = {}
        self._is_populating = False

    def _on_item_changed(self, item):
        if self._is_populating: return
        row = item.row()
        col = item.column()
        if col == 7: # Manual Checkbox changed
            old_item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            old_v = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            
            n_item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            n_v = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            n_code = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            n_has_manual = 1 if self.table.item(row, 7).checkState() == Qt.Checked else 0
            n_manual_text = self.table.item(row, 8).text() if self.table.item(row, 8) else ""
            
            success, msg = self.processor.save_to_db(n_item, n_v, n_code, has_manual=n_has_manual, manual_text=n_manual_text, old_item=old_item, old_v_name=old_v)
            if success:
                self._log("Auto-Saved manual toggle.")
            else:
                self._log(f"Failed to save manual toggle: {msg}")

    def _on_double_click(self, row, col):
        if col in (0, 1, 2, 5, 7): return # Ignore ID, File, Page, Qty, Manual
        
        old_item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
        old_v = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
        
        header = self.table.horizontalHeaderItem(col).text()
        current_val = self.table.item(row, col).text() if self.table.item(row, col) else ""
        
        new_v = None
        if header == "Code" or header == "Manual Text":
            # Show toolbar ONLY for Manual Text
            show_tools = (header == "Manual Text")
            dialog = MultilineInputDialog(self, f"Edit {header}", initial_value=current_val, show_toolbar=show_tools)
            if dialog.exec() == QDialog.Accepted:
                new_v = dialog.result
        else:
            text, ok = QInputDialog.getText(self, "Edit", f"Edit {header}:", text=current_val)
            if ok:
                new_v = text
                
        if new_v is not None:
            self.table.setItem(row, col, QTableWidgetItem(new_v))
            
            n_item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            n_v = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            n_code = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            n_has_manual = 1 if self.table.item(row, 7).checkState() == Qt.Checked else 0
            n_manual_text = self.table.item(row, 8).text() if self.table.item(row, 8) else ""
            
            success, msg = self.processor.save_to_db(n_item, n_v, n_code, has_manual=n_has_manual, manual_text=n_manual_text, old_item=old_item, old_v_name=old_v)
            if success:
                self._log("Auto-Saved mapping update.")
            else:
                QMessageBox.critical(self, "Save Error", f"Failed to save to database.\nError: {msg}")

    def _save_selected_to_db(self):
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges: return
        rows = sorted(list(set([r.topRow() for r in selected_ranges])))
        
        saved_count = 0
        for row in rows:
            item = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            v_name = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            code = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            has_manual = 1 if self.table.item(row, 7).checkState() == Qt.Checked else 0
            manual_text = self.table.item(row, 8).text() if self.table.item(row, 8) else ""
            
            success, msg = self.processor.save_to_db(item, v_name, code, has_manual=has_manual, manual_text=manual_text)
            if success:
                saved_count += 1
            else:
                QMessageBox.critical(self, "Save Error", f"Failed to save {item}.\nError: {msg}")
                return
                
        if saved_count > 0:
            QMessageBox.information(self, "Saved", f"Database updated ({saved_count} items).")

    def _label_pdfs(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No orders in the list to label.")
            return
            
        file_to_pages = defaultdict(lambda: defaultdict(list))
        
        for row in range(self.table.rowCount()):
            code = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            if code == "NOT FOUND" or not code: continue
            
            item_name = self.table.item(row, 3).text() if self.table.item(row, 3) else ""
            variant_name = self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            qty = 1
            try: qty = int(self.table.item(row, 5).text())
            except: pass
            
            has_manual = 1 if self.table.item(row, 7).checkState() == Qt.Checked else 0
            manual_text = self.table.item(row, 8).text() if self.table.item(row, 8) else ""

            if row in self.file_to_order_meta:
                meta = self.file_to_order_meta[row]
                file_to_pages[meta['path']][meta['page']].append({
                    'code': code,
                    'y_pos': meta['y_pos'],
                    'qty': qty,
                    'has_manual': has_manual,
                    'manual_text': manual_text
                })

        if not file_to_pages:
            QMessageBox.warning(self, "No Matches", "No matching codes found in DB to draw.")
            return

        output_files = []
        for file_path, pages_dict in file_to_pages.items():
            fname = os.path.basename(file_path)
            output_name = fname.replace(".pdf", "_labelled.pdf")
            output_path = os.path.join(os.path.dirname(file_path), output_name)
            self._log(f"Drawing Multi-label PDF: {fname}...")
            result_path = self.processor.add_labels_to_pdf(file_path, output_path, pages_dict)
            output_files.append(result_path)

        if output_files:
            QMessageBox.information(self, "Success", f"Labelled {len(output_files)} files.\nList cleared.")
            self._clear_table()
            try:
                if os.name == 'nt': os.startfile(output_files[0])
                else: os.system(f"open '{output_files[0]}'")
            except: pass

if __name__ == "__main__":
    # Prevent scaling issues on Windows
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        
    app = QApplication(sys.argv)
    window = PDFLabelerApp()
    window.show()
    sys.exit(app.exec())