import os
import img2pdf
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import svgwrite
from pdf2docx import Converter

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
        
        # 創建一個新的SVG文件
        dwg = svgwrite.Drawing(svg_file, size=image.size)
        
        # 將圖像轉換為base64編碼
        import io
        import base64
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        # 在SVG中嵌入圖像
        dwg.add(dwg.image(href=f"data:image/png;base64,{img_str}",
                          insert=(0, 0),
                          size=image.size))
        
        # 保存SVG文件
        dwg.save()

def pdf_to_word(pdf_path):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    folder_name = create_output_folder(pdf_path, 'word')
    docx_file = os.path.join(folder_name, f"{base_name}.docx")
    
    # 使用pdf2docx進行轉換
    cv = Converter(pdf_path)
    cv.convert(docx_file)
    cv.close()

def main():
    pdf_path = input("請輸入PDF檔案的路徑: ")
    
    if not os.path.exists(pdf_path):
        print("錯誤：指定的PDF檔案不存在。")
        return
    
    print("請選擇輸出格式:")
    print("1. PNG")
    print("2. JPG")
    print("3. SVG")
    print("4. Word")
    
    choice = input("請輸入您的選擇 (1-4): ")
    
    if choice == '1':
        pdf_to_images(pdf_path, 'png')
        print("轉換為PNG完成。")
    elif choice == '2':
        pdf_to_images(pdf_path, 'jpg')
        print("轉換為JPG完成。")
    elif choice == '3':
        pdf_to_svg(pdf_path)
        print("轉換為SVG完成。")
    elif choice == '4':
        pdf_to_word(pdf_path)
        print("轉換為Word完成。")
    else:
        print("無效的選擇。請重新運行程式並選擇有效的選項。")

if __name__ == "__main__":
    main()