import pdfplumber
import re  
import io
import barcode
from barcode.writer import ImageWriter
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from tkinter import filedialog, Tk
from PIL import Image
import sys
import os
from pylibdmtx.pylibdmtx import encode
import ctypes

dll_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'libdmtx-64.dll')
ctypes.CDLL(dll_path)
def generate_datamatrix_transparent2(data):
    """ 生成透明背景的 Data Matrix 條碼 """
    encoded = encode(data.encode('utf-8'))
    img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
    

    img = img.convert("RGBA")
    datas = img.getdata()
    
 
    new_data = []
    for item in datas:
        if item[:3] == (255, 255, 255):  
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    
    img.putdata(new_data)
    barcode_path = "datamatrix_temp.png"
    img.save(barcode_path)
    return barcode_path


def generate_barcode_transparent(data):
    """ 生成透明背景的一維條碼 (Code 128) """
    code128 = barcode.get_barcode_class('code128')  
    barcode_obj = code128(data, writer=ImageWriter())

    barcode_path = "barcode_temp"  
    
 
    barcode_filename = barcode_obj.save(
        barcode_path,
        options={
            "module_width": 0.2, 
            "module_height": 15,
            "write_text": False
        }
    )


    if not os.path.exists(barcode_filename):
        raise FileNotFoundError(f"{barcode_filename}")


    img = Image.open(barcode_filename).convert("RGBA")
    datas = img.getdata()

    new_data = []
    for item in datas:
        if item[:3] == (255, 255, 255): 
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)

    img.putdata(new_data)


    final_barcode_path = "barcode_temp.png"
    img.save(final_barcode_path)

    if not os.path.exists(final_barcode_path):
        raise FileNotFoundError(f"透明條碼未成功儲存: {final_barcode_path}")

    return final_barcode_path

def process_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf_text:
        pdf_reader = PdfReader(pdf_path)
        writer = PdfWriter()
        pattern = r'工單號碼\s*((?:[A-Z]\d{6,10}|\d{6,10}))'
        
        for i, page in enumerate(pdf_text.pages):
            text = page.extract_text()
            work_order_number = None
            x, y = None, None
            if text:
                match = re.search(pattern, text)
                if match:
                    work_order_number = match.group(1)
                    bbox = page.extract_words()
                    for word in bbox:
                        if word['text'] == work_order_number:
                            x, y = word['x1'], word['top']  
                            break
            
            pdf_page = pdf_reader.pages[i]
            if work_order_number and x is not None and y is not None:
                barcode_string = work_order_number + "-001"
                print(f"第 {i+1} 頁生成的 Data Matrix 內容: {barcode_string}")
                barcode_image_path = generate_datamatrix_transparent2(barcode_string)
                
                page_width = float(pdf_page.mediabox.width)
                page_height = float(pdf_page.mediabox.height)
                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))
                barcode_size = 27  
                margin = 50 
                barcode_x = x + margin  
                barcode_y = page_height - y - 23
                
                can.drawImage(barcode_image_path, barcode_x, barcode_y, width=barcode_size, height=barcode_size, mask='auto')
                can.save()
                packet.seek(0)
                
                overlay_pdf = PdfReader(packet)
                overlay_page = overlay_pdf.pages[0]
                pdf_page.merge_page(overlay_page)
            else:
                print(f"第 {i+1} 頁未找到工單號碼或無法確定位置")
            
            writer.add_page(pdf_page)
        
        with open(pdf_path, 'wb') as f_out:
            writer.write(f_out)
        print(f"✅ PDF 已生成：{pdf_path}")

def process2_pdf(pdf_path):
    """ 在 PDF 中為數量添加條碼 """
    with pdfplumber.open(pdf_path) as pdf_text:
        pdf_reader = PdfReader(pdf_path)
        writer = PdfWriter()
        quantity_pattern = r'數量\s*(\d+)' 
        
        for i, page in enumerate(pdf_text.pages):
            text = page.extract_text()
            quantity = None
            x, y = None, None
            if text:
                quantity_match = re.search(quantity_pattern, text)
                if quantity_match:
                    quantity = quantity_match.group(1)
                bbox = page.extract_words()
                for word in bbox:
                    if quantity and word['text'] == quantity:
                        x, y = word['x1'], word['top']  
                        break
            
            pdf_page = pdf_reader.pages[i]
            if quantity and x is not None and y is not None:
                barcode_string = f"{quantity}"
                print(f"第 {i+1} 頁生成的條碼內容: {barcode_string}")
                barcode_image_path = generate_barcode_transparent(barcode_string)
                
                page_width = float(pdf_page.mediabox.width)
                page_height = float(pdf_page.mediabox.height)
                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))
                barcode_width = 50
                barcode_height = 15  
                margin = 5
                barcode_x = x + margin  
                barcode_y = page_height - y - barcode_height 
                
                can.drawImage(barcode_image_path, barcode_x, barcode_y, width=barcode_width, height=barcode_height, mask='auto')
                can.save()
                packet.seek(0)
                
                overlay_pdf = PdfReader(packet)
                overlay_page = overlay_pdf.pages[0]
                pdf_page.merge_page(overlay_page)
            else:
                print(f"第 {i+1} 頁未找到數量或無法確定位置")
            
            writer.add_page(pdf_page)
        
        with open(pdf_path, 'wb') as f_out:
            writer.write(f_out)
def wait_for_any_key():
    try:
        import msvcrt
        msvcrt.getch()
    except ImportError:
        import termios, tty
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
def select_pdfs():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="選擇 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])
    return file_paths

if __name__ == '__main__':
    pdf_files = list(select_pdfs())
    for pdf_file in pdf_files:
        process2_pdf(pdf_file)
        process_pdf(pdf_file)
    print("請輸入任意鍵結束..")
    wait_for_any_key()