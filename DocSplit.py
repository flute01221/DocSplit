import sys
import os
import tempfile
import shutil
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QVBoxLayout, QHBoxLayout, 
    QLabel, QPushButton, QWidget, QScrollArea, QGridLayout, QComboBox,
    QCheckBox, QSpinBox, QFrame, QMessageBox, QDialog, QGroupBox,
    QRadioButton, QButtonGroup  )

from PySide6.QtGui import QPixmap, QImage
from PySide6.QtCore import Qt, QSize, QThread, Signal
from PySide6.QtGui import QIcon
import fitz 
from pptx import Presentation
import win32com.client  
import comtypes.client
import subprocess
import threading
from PIL import ImageDraw


class ThumbnailWorker(QThread):
    thumbnail_ready = Signal(int, QPixmap)
    finished = Signal()
    
    def __init__(self, file_path, num_pages):
        super().__init__()
        self.file_path = file_path
        self.num_pages = num_pages
        self.is_ppt = file_path.lower().endswith(('.ppt', '.pptx'))
        
    def run(self):
        try:
            if self.is_ppt:
                # 處理PPT的縮圖
                temp_dir = tempfile.mkdtemp()
                ppt_app = None
                presentation = None
                
                try:
                    ppt_app = win32com.client.Dispatch('PowerPoint.Application')                    
                    # 使用絕對路徑
                    abs_file_path = os.path.abspath(self.file_path)                
                    presentation = ppt_app.Presentations.Open(abs_file_path)
                    
                    # 獲取投影片數量
                    slide_count = presentation.Slides.Count
                    
                    for i in range(1, slide_count + 1):
                        temp_path = os.path.join(temp_dir, f"slide_{i}.png")
                        presentation.Slides.Item(i).Export(temp_path, "PNG")
                        
                        if os.path.exists(temp_path):
                            pixmap = QPixmap(temp_path)
                            self.thumbnail_ready.emit(i-1, pixmap)
                except Exception as e:
                    import traceback
                    print(f"PowerPoint處理出錯: {e}\n{traceback.format_exc()}")
                finally:
                    # 釋放資源
                    if presentation:
                        try:
                            presentation.Close()
                        except:
                            pass
                    
                    if ppt_app:
                        try:
                            ppt_app.Quit()
                        except:
                            pass
                    
                    # 確保刪除臨時目錄
                    try:
                        # 文件被使用，稍等一下
                        import time
                        time.sleep(0.5)
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except:
                        pass
            else:
                # 處理PDF的縮圖 (原代碼不變)
                pdf_document = fitz.open(self.file_path)
                for i in range(len(pdf_document)):
                    page = pdf_document.load_page(i)
                    pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                    
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(img)
                    self.thumbnail_ready.emit(i, pixmap)
                
                pdf_document.close()
        except Exception as e:
            import traceback
            print(f"生成縮圖時發生錯誤: {e}\n{traceback.format_exc()}")
        
        self.finished.emit()

class PrintOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("列印選項")
        self.setMinimumWidth(300)
        
        layout = QVBoxLayout()

        # 每頁投影片數量選項
        slides_per_page_group = QGroupBox("每頁投影片數量/Slides per Page")
        slides_layout = QVBoxLayout()

        self.button_group = QButtonGroup(self)
        self.radio_1 = QRadioButton("1張投影片/1 Slide")
        self.radio_1.setChecked(True)
        self.radio_2 = QRadioButton("2張投影片/2 Slides")
        self.radio_4 = QRadioButton("4張投影片/4 Slides")
        self.radio_6 = QRadioButton("6張投影片/6 Slides")
        self.radio_9 = QRadioButton("9張投影片/9 Slides")

        self.button_group.addButton(self.radio_1, 1)
        self.button_group.addButton(self.radio_2, 2)
        self.button_group.addButton(self.radio_4, 4)
        self.button_group.addButton(self.radio_6, 6)
        self.button_group.addButton(self.radio_9, 9)

        slides_layout.addWidget(self.radio_1)
        slides_layout.addWidget(self.radio_2)
        slides_layout.addWidget(self.radio_4)
        slides_layout.addWidget(self.radio_6)
        slides_layout.addWidget(self.radio_9)

        slides_per_page_group.setLayout(slides_layout)
        layout.addWidget(slides_per_page_group)

        # 按鈕
        buttons_layout = QHBoxLayout()
        self.ok_button = QPushButton("確認/OK")
        self.cancel_button = QPushButton("取消/Cancel")

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        buttons_layout.addWidget(self.ok_button)
        buttons_layout.addWidget(self.cancel_button)

        layout.addLayout(buttons_layout)
        self.setLayout(layout)
    
    def get_slides_per_page(self):
        return self.button_group.checkedId() 

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("選擇頁面後重組/Selected Pages Rebuild Document")
        
        self.setMinimumSize(800, 600)
        self.setWindowIcon(QIcon(os.path.abspath("icon.ico")))
        self.file_path = None
        self.thumbnails = []
        self.selected_indexes = []
        self.current_file_type = None
        
        self.init_ui()
    
    def init_ui(self):
        # 主佈局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # 上方按鈕區域
        button_layout = QHBoxLayout()
        
        self.open_button = QPushButton("📂打開檔案/Open File")
        self.open_button.clicked.connect(self.open_file)
        
        self.export_pdf_button = QPushButton("📄匯出為PDF/Export as PDF")
        self.export_pdf_button.clicked.connect(self.export_to_pdf)
        self.export_pdf_button.setEnabled(False)
        
        self.export_ppt_button = QPushButton("✅匯出為PPT/Export as PPT")
        self.export_ppt_button.clicked.connect(self.export_to_ppt)
        self.export_ppt_button.setEnabled(False)

        self.export_word_button = QPushButton("📝匯出為Word/Export as Word")
        self.export_word_button.clicked.connect(self.export_to_word)
        self.export_word_button.setEnabled(False)
        
        self.print_button = QPushButton("📤預覽PDF並列印/Preview PDF and Print")
        self.print_button.clicked.connect(self.print_document)
        self.print_button.setEnabled(False)
        
        button_layout.addWidget(self.open_button)
        button_layout.addWidget(self.export_pdf_button)
        button_layout.addWidget(self.export_word_button)
        button_layout.addWidget(self.export_ppt_button)       
        button_layout.addWidget(self.print_button)
        
        
        main_layout.addLayout(button_layout)
        
        # 縮圖顯示區域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        self.thumbnails_widget = QWidget()
        self.thumbnails_layout = QGridLayout(self.thumbnails_widget)
        self.thumbnails_layout.setSpacing(15) 
        self.thumbnails_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        
        scroll_area.setWidget(self.thumbnails_widget)
        main_layout.addWidget(scroll_area)
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
    
    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇檔案/Select File", "", "文件/Documents (*.ppt *.pptx *.pdf *.docx)"
        )
        
        if not file_path:
            return
            
        self.file_path = file_path
        self.clear_thumbnails()
        self.load_thumbnails()
        
        # 啟用按鈕
        self.export_pdf_button.setEnabled(True)
        self.export_ppt_button.setEnabled(True)
        self.print_button.setEnabled(True)
    
    def clear_thumbnails(self):
        # 清除現有縮圖
        for i in reversed(range(self.thumbnails_layout.count())):
            widget = self.thumbnails_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        self.thumbnails = []
        self.selected_indexes = []
    
    def load_thumbnails(self):
        if not self.file_path:
            return

        ext = os.path.splitext(self.file_path)[1].lower()
        is_word = ext in ['.doc', '.docx']
        # 設置文件類型
        self.current_file_type = 'word' if is_word else ('pdf' if ext == '.pdf' else 'ppt')

        if is_word:
            # Word → PDF
            temp_pdf_path = os.path.join(tempfile.gettempdir(), "word_to_pdf_preview.pdf")
            self.convert_word_to_pdf(self.file_path, temp_pdf_path)
            preview_path = temp_pdf_path
        else:
            preview_path = self.file_path

        # 顯示訊息 + 建立縮圖工作
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("正在生成縮圖...")
        msg.setWindowTitle("處理中")
        msg.setStandardButtons(QMessageBox.NoButton)
        msg.show()
        QApplication.processEvents()

        self.worker = ThumbnailWorker(preview_path, None)
        self.worker.thumbnail_ready.connect(self.add_thumbnail)
        self.worker.finished.connect(msg.close)
        self.worker.start()

    
    def add_thumbnail(self, index, pixmap):
        # 縮放縮圖
        if self.current_file_type == 'word':
        # Word文件使用較大的縮圖
            pixmap = pixmap.scaled(QSize(280, 320), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        else:
            # 其他文件使用標準大小
            pixmap = pixmap.scaled(QSize(200, 150), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        
        # 創建框架包含縮圖
        frame = QFrame()
        frame.setFrameStyle(QFrame.Panel | QFrame.Raised)
        frame.setLineWidth(2)
        
        frame_layout = QVBoxLayout()
        
        # 縮圖標籤
        thumbnail_label = QLabel()
        thumbnail_label.setPixmap(pixmap)
        thumbnail_label.setAlignment(Qt.AlignCenter)
        
        # 頁碼標籤
        page_label = QLabel(f"頁 {index + 1}")
        page_label.setAlignment(Qt.AlignCenter)
        
        frame_layout.addWidget(thumbnail_label)
        frame_layout.addWidget(page_label)
        frame.setLayout(frame_layout)
        
        # 存儲縮圖數據
        self.thumbnails.append({
            'index': index,
            'frame': frame,
            'selected': False
        })
        
        # 添加點擊事件
        frame.mousePressEvent = lambda event, idx=index: self.toggle_selection(idx)
        
        row = index // 4
        col = index % 4
        self.thumbnails_layout.addWidget(frame, row, col)
    
    def toggle_selection(self, index):
        # 切換選擇狀態
        for thumbnail in self.thumbnails:
            if thumbnail['index'] == index:
                thumbnail['selected'] = not thumbnail['selected']
                
                if thumbnail['selected']:
                    thumbnail['frame'].setStyleSheet("background-color: #e0e0ff;")
                    if index not in self.selected_indexes:
                        self.selected_indexes.append(index)
                else:
                    thumbnail['frame'].setStyleSheet("")
                    if index in self.selected_indexes:
                        self.selected_indexes.remove(index)
                break
    
    def export_to_pdf(self):
        """將選定頁面匯出為PDF"""
        if not self.file_path:
            return
                
        if not self.selected_indexes:
            QMessageBox.warning(self, "Warning", "Please select pages to export first")
            return
                
        # 選擇保存位置
        save_path, _ = QFileDialog.getSaveFileName(self, "儲存/Save PDF", "", "PDF (*.pdf)")
        if not save_path:
            return
            
        try:
            # 確保文件名有 .pdf 副檔名
            if not save_path.lower().endswith('.pdf'):
                save_path += '.pdf'
            
            # 使用絕對路徑
            abs_file_path = os.path.abspath(self.file_path)
            abs_save_path = os.path.abspath(save_path)
            
            # 檢查是PDF還是PPT
            is_pdf = self.file_path.lower().endswith('.pdf')
            is_ppt = self.file_path.lower().endswith(('.ppt', '.pptx'))
            is_word = self.file_path.lower().endswith(('.doc', '.docx'))
            
            if is_pdf:
                # PDF到PDF的處理
                try:
                    # 打開PDF
                    pdf_document = fitz.open(abs_file_path)
                    # 創建新的PDF
                    new_pdf = fitz.open()
                    
                    # 複製選定的頁面
                    for idx in sorted(self.selected_indexes):
                        new_pdf.insert_pdf(pdf_document, from_page=idx, to_page=idx)
                    
                    # 保存新PDF
                    new_pdf.save(abs_save_path)
                    new_pdf.close()
                    pdf_document.close()
                    
                    QMessageBox.information(self, "Success", "PDF exported successfully!")
                    
                except Exception as e:
                    import traceback
                    error_msg = f"處理PDF時發生錯誤/An error occurred while processing the PDF:\n{str(e)}\n\n{traceback.format_exc()}"
                    QMessageBox.critical(self, "Error", error_msg)
                    print(error_msg)
                    
            else:
                # PPT到PDF的處理
                ppt_app = None
                presentation = None
                temp_presentation = None
                
                try:
                    # 使用win32com創建PowerPoint
                    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                    
                    # 打開原始文件
                    presentation = ppt_app.Presentations.Open(abs_file_path)
                    
                    # 創建臨時投影片
                    temp_presentation = ppt_app.Presentations.Add()
                    
                    # 複製選定的投影片
                    for idx in sorted(self.selected_indexes):
                        slide_index = idx + 1
                        presentation.Slides.Item(slide_index).Copy()
                        temp_presentation.Slides.Paste()
                    
                    # 保存為PDF
                    temp_presentation.SaveAs(abs_save_path, 32)
                    
                    QMessageBox.information(self, "Success", "PDF exported successfully!")
                    
                finally:                    
                    if temp_presentation:
                        try:
                            temp_presentation.Close()
                        except:
                            pass                            
                    if presentation:
                        try:
                            presentation.Close()
                        except:
                            pass                            
                    if ppt_app:
                        try:
                            ppt_app.Quit()
                        except:
                            pass
            
        except Exception as e:
            import traceback
            error_msg = f"An error occurred while processing the PDF:\n{str(e)}\n\n{traceback.format_exc()}"
            QMessageBox.critical(self, "Error", error_msg)
            print(error_msg) 
   
    def convert_word_to_pdf(self, docx_path, pdf_path):
        try:
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            word.Quit()
        except Exception as e:
            import traceback
            error_msg = f"從Word轉換到PDF時發生錯誤: {str(e)}\n\n{traceback.format_exc()}"
            QMessageBox.critical(self, "錯誤", error_msg)
            print(error_msg)

    def export_to_ppt(self):
        """將選定頁面匯出為PPT"""
        if not self.file_path:
            return
                
        if not self.selected_indexes:
            QMessageBox.warning(self, "Warning", "請先選擇要匯出的頁面/Please select pages to export first")
            return
                
        # 選擇保存位置
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存PPT", "", "PowerPoint (*.pptx)"
        )
        
        if not save_path:
            return
            
        try:
            # 確保副檔名是 .pptx
            if not save_path.lower().endswith('.pptx'):
                save_path += '.pptx'
            
            ppt_app = None
            presentation = None
            new_presentation = None
            
            try:
                ppt_app = win32com.client.Dispatch('PowerPoint.Application')
                
                # 根據源文件類型處理
                is_ppt = self.file_path.lower().endswith(('.ppt', '.pptx'))
                is_pdf = self.file_path.lower().endswith('.pdf')
                is_word = self.file_path.lower().endswith(('.doc', '.docx'))

                if is_ppt:
                    # 使用絕對路徑
                    abs_file_path = os.path.abspath(self.file_path)
                    abs_save_path = os.path.abspath(save_path)                    
                    # 打開原始文件
                    presentation = ppt_app.Presentations.Open(abs_file_path)                    
                    # 創建新投影片
                    new_presentation = ppt_app.Presentations.Add()    
                    # 複製選取的投影片                
                    for idx in sorted(self.selected_indexes):
                        # 使用 Item 方法
                        slide_index = idx + 1
                        presentation.Slides.Item(slide_index).Copy()                        
                        new_presentation.Slides.Paste()
                    
                    # 保存
                    new_presentation.SaveAs(abs_save_path)
                    
                    QMessageBox.information(self, "Success", "Exported successfully!")
                else:
                    pass
                    
            finally:
              
                if new_presentation:
                    try:
                        new_presentation.Close()
                    except:
                        pass
                        
                if presentation:
                    try:
                        presentation.Close()
                    except:
                        pass
                        
                if ppt_app:
                    try:
                        ppt_app.Quit()
                    except:
                        pass
            
        except Exception as e:
            import traceback
            error_msg = f"無法匯出PPT:\n{str(e)}\n\n{traceback.format_exc()}"
            QMessageBox.critical(self, "錯誤", error_msg)
            print(error_msg)  

    def export_to_word(self):
        if not self.file_path:
            return

        if not self.selected_indexes:
            QMessageBox.warning(self, "Warning", "Please select pages to export first")
            return

        try:
            QMessageBox.information(self, "尚未完成/Not Yet Implemented", "Word 匯出功能尚未完成/Word export feature is not yet available!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"無法匯出/Unable to export Word：{e}")


    def print_document(self):
        """列印選定頁面"""
        if not self.file_path:
            return
                
        if not self.selected_indexes:
            QMessageBox.warning(self, "警告", "請先選擇要列印的頁面")
            return
        
        # 顯示列印選項對話框
        print_dialog = PrintOptionsDialog(self)
        result = print_dialog.exec_()
        
        if result != QDialog.Accepted:
            return
                
        slides_per_page = print_dialog.get_slides_per_page()
        
        try:
            # 根據源文件類型處理
            is_ppt = self.file_path.lower().endswith(('.ppt', '.pptx'))
            is_pdf = self.file_path.lower().endswith('.pdf')          
            is_word = self.file_path.lower().endswith(('.doc', '.docx'))

            if is_ppt:
                # 從PPT列印
                temp_dir = tempfile.mkdtemp()
                temp_pdf = os.path.join(temp_dir, "temp_print.pdf")
                
                try:
                    # 使用 PowerPoint 創建 PDF
                    ppt_app = win32com.client.Dispatch('PowerPoint.Application')                  
                    presentation = ppt_app.Presentations.Open(os.path.abspath(self.file_path))
                    new_presentation = ppt_app.Presentations.Add()
                    
                    # 複製選定的投影片
                    for idx in sorted(self.selected_indexes):
                        presentation.Slides.Item(idx + 1).Copy()
                        new_presentation.Slides.Paste()
                    
                    # 保存為PDF
                    new_presentation.SaveAs(os.path.abspath(temp_pdf), 32)
                    
                    # 關閉投影片
                    new_presentation.Close()
                    presentation.Close()
                    ppt_app.Quit()
                    
                    # 現在列印生成的 PDF 文件
                    # 使用適當的選項設置
                    pdf_document = fitz.open(temp_pdf)
                    
                    # 創建適合列印的新 PDF
                    print_pdf = os.path.join(temp_dir, "print_ready.pdf")
                    
                    if slides_per_page == 1:
                        # 直接列印，不需要特殊處理
                        pdf_document.save(print_pdf)
                    else:
                        # 使用 PyMuPDF 創建多投影片每頁的版本
                        doc_out = fitz.open()
                        page_width, page_height = fitz.paper_size("a4")
                        
                        # 根據每頁投影片數量計算布局
                        if slides_per_page == 2:
                            # 2張投影片每頁，縱向排列
                            rows, cols = 2, 1
                        elif slides_per_page == 4:
                            # 4張投影片每頁，2x2 網格
                            rows, cols = 2, 2
                        elif slides_per_page == 6:
                            # 6張投影片每頁，3x2 網格
                            rows, cols = 3, 2
                        elif slides_per_page == 9:
                            # 9張投影片每頁，3x3 網格
                            rows, cols = 3, 3
                        else:
                            # 默認使用 1 張投影片每頁
                            rows, cols = 1, 1
                        
                        # 計算每張投影片的尺寸
                        cell_width = page_width / cols
                        cell_height = page_height / rows
                        
                        # 計算需要多少頁
                        pages_count = pdf_document.page_count
                        output_pages = (pages_count + slides_per_page - 1) // slides_per_page
                        
                        # 創建輸出頁面
                        for p in range(output_pages):
                            page_out = doc_out.new_page(width=page_width, height=page_height)
                            
                            # 添加當前頁的投影片
                            for i in range(slides_per_page):
                                idx = p * slides_per_page + i
                                if idx >= pages_count:
                                    break
                                
                                # 計算投影片位置
                                row = i // cols
                                col = i % cols
                                
                                # 計算目標矩形
                                target_rect = fitz.Rect(
                                    col * cell_width,
                                    row * cell_height,
                                    (col + 1) * cell_width,
                                    (row + 1) * cell_height
                                )
                                
                                # 添加投影片到頁面
                                page_in = pdf_document[idx]
                                page_out.show_pdf_page(target_rect, pdf_document, idx)
                        
                        # 保存文件
                        doc_out.save(print_pdf)
                        doc_out.close()
                    
                    pdf_document.close()                    

                    # 改為讓使用者自行開啟後列印
                    if os.name == 'nt':
                        os.startfile(print_pdf)  
                    else:
                        subprocess.call(['xdg-open', print_pdf])

                    QMessageBox.information(self, "預覽列印/Print Preview", "已開啟 PDF/ PDF opened.，請在檢視器中使用列印功能）。")
                    
                except Exception as e:
                    import traceback
                    error_msg = f"處理列印時出錯/An error occurred while printing: {e}\n{traceback.format_exc()}"
                    QMessageBox.critical(self, "Error", error_msg)
                    print(error_msg)
                finally:
                    # 刪除臨時文件
                    try:
                        import time
                        time.sleep(1) 
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except:
                        pass
            else:
                # 創建臨時PDF
                temp_dir = tempfile.mkdtemp()
                temp_pdf = os.path.join(temp_dir, "temp_print.pdf")
                
                try:
                    pdf_document = fitz.open(self.file_path)
                    new_pdf = fitz.open()
                    
                    for idx in sorted(self.selected_indexes):
                        new_pdf.insert_pdf(pdf_document, from_page=idx, to_page=idx)
                    
                    # 處理每頁多張投影片的設置
                    if slides_per_page > 1:
                        # 使用 PyMuPDF 創建多投影片每頁的版本
                        print_pdf = os.path.join(temp_dir, "print_ready.pdf")
                        doc_out = fitz.open()
                        page_width, page_height = fitz.paper_size("a4")
                        
                        # 根據每頁投影片數量計算布局
                        if slides_per_page == 2:
                            rows, cols = 2, 1
                        elif slides_per_page == 4:
                            rows, cols = 2, 2
                        elif slides_per_page == 6:
                            rows, cols = 3, 2
                        elif slides_per_page == 9:
                            rows, cols = 3, 3
                        else:
                            rows, cols = 1, 1
                        
                        # 計算每張投影片的尺寸
                        cell_width = page_width / cols
                        cell_height = page_height / rows
                        
                        # 計算需要多少頁
                        pages_count = new_pdf.page_count
                        output_pages = (pages_count + slides_per_page - 1) // slides_per_page
                        
                        # 創建輸出頁面
                        for p in range(output_pages):
                            page_out = doc_out.new_page(width=page_width, height=page_height)
                            
                            # 添加當前頁的投影片
                            for i in range(slides_per_page):
                                idx = p * slides_per_page + i
                                if idx >= pages_count:
                                    break
                                
                                # 計算投影片位置
                                row = i // cols
                                col = i % cols
                                
                                # 計算目標矩形
                                target_rect = fitz.Rect(
                                    col * cell_width,
                                    row * cell_height,
                                    (col + 1) * cell_width,
                                    (row + 1) * cell_height
                                )
                                
                                # 添加投影片到頁面
                                page_out.show_pdf_page(target_rect, new_pdf, idx)
                        
                        # 保存文件
                        doc_out.save(print_pdf)
                        doc_out.close()
                    else:
                        # 單張投影片每頁
                        new_pdf.save(temp_pdf)
                        print_pdf = temp_pdf
                    
                    new_pdf.close()
                    pdf_document.close()
                    
                    # 使用系統默認PDF查看器列印
                    if os.name == 'nt':  # Windows
                        os.startfile(print_pdf, 'print')
                    else:
                        # Linux或Mac
                        subprocess.call(['xdg-open', print_pdf])
                    
                    # 等待用戶完成列印
                    QMessageBox.information(self, "列印/Print", "文件已發送到列印機。請在完成後點擊確認。The document has been sent to the printer. Please click OK when done.")
                    
                except Exception as e:
                    import traceback
                    error_msg = f"處理列印時出錯: {e}\n{traceback.format_exc()}"
                    QMessageBox.critical(self, "錯誤", error_msg)
                    print(error_msg)
                finally:
                    # 刪除臨時文件
                    try:
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except:
                        pass
        
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"無法列印: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # ✨ 在這裡加上美化樣式
    app.setStyleSheet("""
        QWidget {
            background-color: #f7f9fc;
            font-family: 'Segoe UI', 'Noto Sans TC', sans-serif;
            font-size: 14px;
            color: #333333;
        }

        QPushButton {
            background-color: #4e8cff;
            color: white;
            padding: 6px 14px;
            border: none;
            border-radius: 6px;
        }
        QPushButton:hover {
            background-color: #3b6eea;
        }
        QPushButton:disabled {
            background-color: #b0c3e6;
        }

        QFrame {
            border: 1px solid #d0d7e4;
            border-radius: 6px;
            background-color: white;
        }

        QLabel {
            padding: 4px;
        }

        QScrollArea {
            border: none;
        }

        QCheckBox {
            spacing: 8px;
        }

        QComboBox {
            background-color: white;
            border: 1px solid #ccc;
            padding: 4px;
            border-radius: 4px;
        }
    """)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())