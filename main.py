

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdfminer.high_level import extract_text
from ebooklib import epub
import re
import fitz

def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
       
        text += page.get_text("text") + "\n\n"  # Thêm \n\n giữa các trang
    return text



# Hàm chuyển đổi PDF sang EPUB với vòng loading
def convert_pdf_to_epub(pdf_path, selected_font, progress_bar):
    try:
        # Bắt đầu vòng loading
        progress_bar['value'] = 0
        root.update_idletasks()

        # Trích xuất nội dung từ PDF (20%)
        text = extract_text(pdf_path)
        if not text.strip():
            raise ValueError("Không thể trích xuất nội dung từ file PDF.")
        progress_bar['value'] = 20
        root.update_idletasks()
        # Xuất văn bản trung gian ra file .txt để kiểm tra (30%)
        output_dir = os.path.join(os.path.expanduser("~"), "Documents", "epub")
        os.makedirs(output_dir, exist_ok=True)  # Tạo thư mục nếu chưa tồn tại
        txt_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_path))[0] + "_raw.txt")
        

        # Lưu văn bản gốc với các ký tự xuống dòng hiển thị
        text_with_visible_newlines = text.replace('\n', '/n')
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(text_with_visible_newlines)
        print(f"Đã xuất file văn bản trung gian: {txt_path}")
        progress_bar['value'] = 30
        root.update_idletasks()

        # Bước 1: Xử lý tiêu đề đặc biệt
        text = re.sub(
            r'/n\s+/n(Giảng giải|Phần|Chương|Bài|Mục)\s+([^/]+?)(/n\s+/n|\Z)',
            r'</p><h2>\1 \2</h2><p>',
            text_with_visible_newlines,
            flags=re.IGNORECASE
        )

        # Bước 2: Xử lý các đoạn xuống dòng kép (/n /n) - đây là ngắt đoạn thực sự
        text = re.sub(r'/n\s+/n', '</p><p>', text)

        # Bước 3: Xử lý các ngắt dòng đơn (/n)
        # Chỉ tạo ngắt đoạn nếu dòng kết thúc bằng dấu câu và theo sau là khoảng trắng
        text = re.sub(r'([.!?])\s+/n\s+', r'\1</p><p>', text)  # Dấu câu + khoảng trắng + /n + khoảng trắng
        text = re.sub(r'([.!?])/n\s+', r'\1</p><p>', text)      # Dấu câu + /n + khoảng trắng

        # Các trường hợp ngắt dòng đơn khác (không theo quy tắc trên) - thay thế bằng khoảng trắng
        text = re.sub(r'/n', ' ', text)

        # Bước 4: Xử lý khoảng trắng thừa
        text = re.sub(r'\s+', ' ', text)

        # Bước 5: Xử lý ngắt đoạn theo dấu câu trong văn bản
        # Chỉ tạo ngắt đoạn nếu dấu câu theo sau bởi khoảng trắng (không phải là viết tắt như "Dr.")
        # # Tìm kiếm dấu câu và theo sau bởi ít nhất một khoảng trắng
        # text = re.sub(r'([.!?])\s+(?!<p>|[A-Z][a-z]\.)', r'\1</p><p>', text)
         # Tìm kiếm dấu câu và theo sau bởi chính xác 5 khoảng trắng
        text = re.sub(r'([.!?)\s{5}(?!<p>|[A-Z][a-z]\.)', r'\1</p><p>', text)

        # Bước 6: Đảm bảo bắt đầu và kết thúc với thẻ HTML phù hợp
        if not text.startswith('<p>') and not text.startswith('<h2>'):
            text = f'<p>{text}'
        if not text.endswith('</p>'):
            text = f'{text}</p>'

        # Bước 7: Dọn dẹp thẻ HTML
        text = re.sub(r'<p>\s*</p>', '', text)  # Loại bỏ thẻ p rỗng
        text = re.sub(r'</p>\s*<p>', '</p><p>', text)  # Chuẩn hóa khoảng cách
        text = re.sub(r'</h2><p></p>', '</h2><p>', text)  # Loại bỏ thẻ p rỗng sau h2
        text = re.sub(r'<p>\s*</h2>', '</h2>', text)  # Sửa thẻ p trước h2

        # Tạo sách EPUB
        book = epub.EpubBook()
        book.set_identifier('id123456')
        book.set_title(os.path.basename(pdf_path).replace('.pdf', ''))
        book.set_language('vi')
        book.add_author('Tác giả')
        progress_bar['value'] = 40
        root.update_idletasks()

        # Tạo chương đầu tiên
        chapter = epub.EpubHtml(title='Chương 1', file_name='chap_01.xhtml', lang='vi')
        html_content = f'<h1>Chương 1</h1><p>{text}</p>'
        chapter.content = html_content
        book.add_item(chapter)
        progress_bar['value'] = 60
        root.update_idletasks()

        # Thêm CSS với font được chọn (80%)
        css_content = f'''
        @font-face {{
            font-family: "{selected_font}";
            src: url("fonts/{selected_font}-Regular.ttf");
        }}
       
        body {{
            font-family: "{selected_font}", sans-serif;
            line-height: 1.5;
            font-size: 1em;
        }}
        h1, h2, h3, h4, h5, h6 {{
            font-weight: bold;
        }}
        '''
        style = epub.EpubItem(uid="style_nav", file_name="styles/style.css", media_type="text/css", content=css_content)
        book.add_item(style)

        # Kiểm tra font và nhúng vào EPUB
        font_path = os.path.join(os.path.dirname(__file__), "fonts", selected_font, "static", f"{selected_font}-Regular.ttf")
        if os.path.exists(font_path):
            with open(font_path, 'rb') as f:
                font_data = f.read()
            font_item = epub.EpubItem(uid=f"font_{selected_font.lower()}", file_name=f"fonts/{selected_font}-Regular.ttf",
                                      media_type="application/x-font-ttf", content=font_data)
            book.add_item(font_item)
        else:
            print(f"Không tìm thấy font {selected_font}, tiếp tục mà không nhúng font.")
        progress_bar['value'] = 80
        root.update_idletasks()

        # Cấu hình EPUB và lưu file (100%)
        book.spine = ['nav', chapter]
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())

        output_dir = os.path.join(os.path.expanduser("~"), "Documents", "epub")
        os.makedirs(output_dir, exist_ok=True)  # Tạo thư mục nếu chưa tồn tại
        output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".epub")
        epub.write_epub(output_path, book)
        progress_bar['value'] = 100
        root.update_idletasks()

        # Hiển thị thông báo thành công
        messagebox.showinfo("Thành công", f"File EPUB đã được tạo thành công:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")

# Giao diện chính
def create_gui():
    def select_file():
        # Mở hộp thoại chọn file PDF
        pdf_path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if pdf_path:
            # Hiển thị Progressbar
            progress_bar.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
            # Thực hiện chuyển đổi với font được chọn
            convert_pdf_to_epub(pdf_path, font_var.get(), progress_bar)
            # Ẩn Progressbar sau khi hoàn thành
            progress_bar.grid_remove()

    # Tạo cửa sổ chính
    global root
    root = tk.Tk()
    root.title("Chuyển đổi PDF sang EPUB")

    # Nhãn hướng dẫn
    tk.Label(root, text="Chọn file PDF để chuyển đổi sang EPUB:", padx=10, pady=10).grid(row=0, column=0, columnspan=2)

    # Dropdown chọn font
    tk.Label(root, text="Chọn font:").grid(row=1, column=0, pady=5, sticky="e")
    fonts = ["Montserrat", "Nunito", "OpenSans", "Pacifico"]
    font_var = tk.StringVar(root)
    font_var.set(fonts[0])  # Giá trị mặc định
    font_menu = tk.OptionMenu(root, font_var, *fonts)
    font_menu.grid(row=1, column=1, pady=5, sticky="w")

    # Nút chọn file
    tk.Button(root, text="Chọn file PDF", command=select_file).grid(row=2, column=0, columnspan=2, pady=10)

    # Progressbar (ẩn ban đầu)
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
    progress_bar.grid_remove()  # Ẩn Progressbar ban đầu

    # Chạy giao diện
    root.mainloop()

if __name__ == "__main__":
    create_gui()