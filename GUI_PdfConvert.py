import PySimpleGUI as sg
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
    return folder_name

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
    return folder_name

def pdf_to_word(pdf_path):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    folder_name = create_output_folder(pdf_path, 'word')
    docx_file = os.path.join(folder_name, f"{base_name}.docx")
    
    cv = Converter(pdf_path)
    cv.convert(docx_file)
    cv.close()
    return folder_name

def main():
    sg.theme('LightBlue2')

    layout = [
        [sg.Text('拖放PDF文件到這裡或點擊瀏覽：')],
        [sg.Input(key='-FILE-'), sg.FileBrowse('瀏覽', file_types=(("PDF Files", "*.pdf"),))],
        [sg.Text('選擇輸出格式：')],
        [sg.Radio('PNG', 'FORMAT', key='-PNG-'), sg.Radio('JPG', 'FORMAT', key='-JPG-'),
         sg.Radio('SVG', 'FORMAT', key='-SVG-'), sg.Radio('Word', 'FORMAT', key='-WORD-')],
        [sg.Button('轉換'), sg.Button('退出')],
        [sg.Output(size=(60, 10))]
    ]

    window = sg.Window('PDF轉換工具', layout, finalize=True)
    window['-FILE-'].bind("<Drop>", "_drop")

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == '退出':
            break
        elif event == '-FILE-' + '_drop':
            window['-FILE-'].update(values['-FILE-' + '_drop'][0])
        elif event == '轉換':
            pdf_path = values['-FILE-']
            if not pdf_path or not os.path.exists(pdf_path):
                print("錯誤：請先選擇一個有效的PDF文件。")
                continue

            if values['-PNG-']:
                output_folder = pdf_to_images(pdf_path, 'png')
                print(f"轉換為PNG完成。文件保存在：{output_folder}")
            elif values['-JPG-']:
                output_folder = pdf_to_images(pdf_path, 'jpg')
                print(f"轉換為JPG完成。文件保存在：{output_folder}")
            elif values['-SVG-']:
                output_folder = pdf_to_svg(pdf_path)
                print(f"轉換為SVG完成。文件保存在：{output_folder}")
            elif values['-WORD-']:
                output_folder = pdf_to_word(pdf_path)
                print(f"轉換為Word完成。文件保存在：{output_folder}")
            else:
                print("錯誤：請選擇一個輸出格式。")

    window.close()

if __name__ == "__main__":
    main()