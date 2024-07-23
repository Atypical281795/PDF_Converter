import os
import sys
import img2pdf
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import svgwrite
from pdf2docx import Converter
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

class PDFConverterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF轉換器")
        self.setGeometry(100, 100, 400, 300)
        self.pdf_path = ""

        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        layout = QVBoxLayout()

        self.drop_label = QLabel("將PDF文件拖放到這裡或按'瀏覽'按鈕")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet("border: 2px dashed #aaa; padding: 20px;")
        self.drop_label.setAcceptDrops(True)
        self.drop_label.dragEnterEvent = self.dragEnterEvent
        self.drop_label.dropEvent = self.dropEvent

        layout.addWidget(self.drop_label)

        button_layout = QHBoxLayout()
        self.select_button = QPushButton("瀏覽")
        self.png_button = QPushButton("轉換為PNG")
        self.jpg_button = QPushButton("轉換為JPG")
        self.svg_button = QPushButton("轉換為SVG")
        self.word_button = QPushButton("轉換為Word")

        self.select_button.clicked.connect(self.open_file_dialog)
        self.png_button.clicked.connect(lambda: self.convert_pdf('png'))
        self.jpg_button.clicked.connect(lambda: self.convert_pdf('jpg'))
        self.svg_button.clicked.connect(lambda: self.convert_pdf('svg'))
        self.word_button.clicked.connect(lambda: self.convert_pdf('word'))

        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.png_button)
        button_layout.addWidget(self.jpg_button)
        button_layout.addWidget(self.svg_button)
        button_layout.addWidget(self.word_button)

        layout.addLayout(button_layout)

        main_widget.setLayout(layout)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files and files[0].lower().endswith('.pdf'):
            self.pdf_path = files[0]
            self.drop_label.setText(f"已選擇文件: {os.path.basename(self.pdf_path)}")
        else:
            QMessageBox.warning(self, "錯誤", "請選擇一個有效的PDF文件。")

    def open_file_dialog(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "選擇PDF文件", "", "PDF Files (*.pdf);;All Files (*)", options=options)
        if file_path:
            self.pdf_path = file_path
            self.drop_label.setText(f"已選擇文件: {os.path.basename(self.pdf_path)}")

    def convert_pdf(self, output_format):
        if not self.pdf_path:
            QMessageBox.warning(self, "錯誤", "請先選擇一個PDF文件。")
            return

        try:
            if output_format in ['png', 'jpg']:
                pdf_to_images(self.pdf_path, output_format)
            elif output_format == 'svg':
                pdf_to_svg(self.pdf_path)
            elif output_format == 'word':
                pdf_to_word(self.pdf_path)
            
            QMessageBox.information(self, "成功", f"轉換為{output_format.upper()}完成。")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"轉換過程中發生錯誤：{str(e)}")

def create_output_folder(pdf_path, output_format):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    folder_name = f"{output_format}_{base_name}"
    os.makedirs(folder_name, exist_ok=True)
    return folder_name

def pdf_to_images(pdf_path, output_format):
    images = convert_from_path(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    folder_name = create_output_folder(pdf_path, output_format)
    
    for i, image in enumerate(images):
        if output_format.lower() == 'png':
            image.save(os.path.join(folder_name, f"{base_name}_page_{i+1}.png"), 'PNG')
        elif output_format.lower() == 'jpg':
            image.save(os.path.join(folder_name, f"{base_name}_page_{i+1}.jpg"), 'JPEG')

def pdf_to_svg(pdf_path):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    images = convert_from_path(pdf_path)
    folder_name = create_output_folder(pdf_path, 'svg')
    
    for i, image in enumerate(images):
        svg_file = os.path.join(folder_name, f"{base_name}_page_{i+1}.svg")
        
        dwg = svgwrite.Drawing(svg_file, size=image.size)
        
        import io
        import base64
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        dwg.add(dwg.image(href=f"data:image/png;base64,{img_str}",
                          insert=(0, 0),
                          size=image.size))
        
        dwg.save()

def pdf_to_word(pdf_path):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    folder_name = create_output_folder(pdf_path, 'word')
    docx_file = os.path.join(folder_name, f"{base_name}.docx")
    
    cv = Converter(pdf_path)
    cv.convert(docx_file)
    cv.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFConverterGUI()
    window.show()
    sys.exit(app.exec_())
