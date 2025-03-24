

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdfminer.high_level import extract_text
from ebooklib import epub
import re
import fitz
import pdfplumber


# def extract_text(pdf_path):
#     doc = fitz.open(pdf_path)
#     text = ""
#     for page in doc:
#         text += page.get_text("text") + "\n\n"  # Thêm \n\n giữa các trang
#     return text

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
        text_with_visible_newlines = re.sub(
            r'/n\s+/n(Giảng giải|Phần|Chương|Bài|Mục)\s+([^/]+?)(/n\s+/n|\Z)',
            r'</p><h2>\1 \2</h2><p>',
            text_with_visible_newlines,
            flags=re.IGNORECASE
        )

        # Bước 2: Xử lý các đoạn xuống dòng kép (/n /n) - đây là ngắt đoạn thực sự
        text_with_visible_newlines = re.sub(r'/n\s*/n', '</p><p>', text_with_visible_newlines)

        # Bước 3: Xử lý các ngắt dòng đơn (/n)
        text_with_visible_newlines = re.sub(r'([.!?])\s+/n\s+', r'\1</p><p>', text_with_visible_newlines)
        text_with_visible_newlines = re.sub(r'([.!?])/n\s+', r'\1</p><p>', text_with_visible_newlines)

        # Các trường hợp ngắt dòng đơn khác (không theo quy tắc trên) - thay thế bằng khoảng trắng
        text_with_visible_newlines = re.sub(r'/n', ' ', text_with_visible_newlines)

        # Bước 4: Xử lý khoảng trắng thừa
        text_with_visible_newlines = re.sub(r'\s+', ' ', text_with_visible_newlines)

        # Bước 6: Đảm bảo bắt đầu và kết thúc với thẻ HTML phù hợp
        if not text_with_visible_newlines.startswith('<p>') and not text_with_visible_newlines.startswith('<h2>'):
            text_with_visible_newlines = f'<p>{text_with_visible_newlines}'
        if not text_with_visible_newlines.endswith('</p>'):
            text_with_visible_newlines = f'{text_with_visible_newlines}</p>'

        # Bước 7: Dọn dẹp thẻ HTML
        text_with_visible_newlines = re.sub(r'<p>\s*</p>', '', text_with_visible_newlines)
        text_with_visible_newlines = re.sub(r'</p>\s*<p>', '</p><p>', text_with_visible_newlines)
        text_with_visible_newlines = re.sub(r'</h2><p></p>', '</h2><p>', text_with_visible_newlines)
        text_with_visible_newlines = re.sub(r'<p>\s*</h2>', '</h2>', text_with_visible_newlines)

        # Tạo sách EPUB
        book = epub.EpubBook()  # Ensure book is initialized
        book.set_identifier('id123456')
        book.set_title(os.path.basename(pdf_path).replace('.pdf', ''))
        book.set_language('vi')
        book.add_author('Tác giả')
        progress_bar['value'] = 40
        root.update_idletasks()

        # Tạo chương đầu tiên
        chapter = epub.EpubHtml(title='Wish you happy!', file_name='chap_01.xhtml', lang='vi')
        html_content = f'<h1>Wish you happy!</h1><p>{text_with_visible_newlines}</p>'
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

# # Giao diện chính
# def create_gui():
#     def select_file():
#         # Mở hộp thoại chọn file PDF
#         pdf_path = filedialog.askopenfilename(
#             title="Chọn file PDF",
#             filetypes=[("PDF files", "*.pdf")]
#         )
#         if pdf_path:
#             # Hiển thị Progressbar
#             progress_bar.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
#             # Thực hiện chuyển đổi với font được chọn
#             convert_pdf_to_epub(pdf_path, font_var.get(), progress_bar)
#             # Ẩn Progressbar sau khi hoàn thành
#             progress_bar.grid_remove()

#     # Tạo cửa sổ chính
#     global root
#     root = tk.Tk()
#     root.title("Chuyển đổi PDF sang EPUB")

#     # Nhãn hướng dẫn
#     tk.Label(root, text="Chọn file PDF để chuyển đổi sang EPUB:", padx=10, pady=10).grid(row=0, column=0, columnspan=2)

#     # Dropdown chọn font
#     tk.Label(root, text="Chọn font:").grid(row=1, column=0, pady=5, sticky="e")
#     fonts = ["Montserrat", "Nunito", "OpenSans", "Pacifico"]
#     font_var = tk.StringVar(root)
#     font_var.set(fonts[0])  # Giá trị mặc định
#     font_menu = tk.OptionMenu(root, font_var, *fonts)
#     font_menu.grid(row=1, column=1, pady=5, sticky="w")

#     # Nút chọn file
#     tk.Button(root, text="Chọn file PDF", command=select_file).grid(row=2, column=0, columnspan=2, pady=10)

#     # Progressbar (ẩn ban đầu)
#     progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
#     progress_bar.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")
#     progress_bar.grid_remove()  # Ẩn Progressbar ban đầu

#     # Chạy giao diện
#     root.mainloop()
# Giao diện chính
# Giao diện chính
def create_gui():
    def select_file():
        # Mở hộp thoại chọn file PDF
        pdf_path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if pdf_path:
            file_label.config(text=os.path.basename(pdf_path))
            status_label.config(text="Đang chuyển đổi...")
            
            # Hiển thị progress bar
            progress_bar.pack(pady=10, padx=30, fill="x")
            
            # Thực hiện chuyển đổi với font được chọn
            convert_pdf_to_epub(pdf_path, font_var.get(), progress_bar)
            
            # Cập nhật trạng thái sau khi hoàn thành
            status_label.config(text="Chuyển đổi thành công!")


    def convert_file():
        # Thực hiện chuyển đổi (giả lập)
        status_label.config(text="Đang chuyển đổi...")
        root.after(2000, lambda: status_label.config(text="Chuyển đổi thành công!"))

    # Tạo cửa sổ chính
    global root  # Make root global so it can be accessed in convert_pdf_to_epub
    root = tk.Tk()
    root.title("PDF to EPUB Converter")
    root.geometry("400x450")  # Increased height for progress bar
    root.configure(bg="#f5f5f5")

    # Thiết lập style cho ttk
    style = ttk.Style()
    style.configure("TFrame", background="#f5f5f5")
    style.configure("TButton", background="#2a7e85", foreground="white", font=("OpenSans", 10))
    style.configure("TCombobox", background="white", foreground="#333333")
    style.map("TButton", background=[("active", "#236b70")])

    # Tiêu đề
    title_label = tk.Label(
        root, 
        text="PDF to EPUB Converter", 
        font=("OpenSans", 18, "bold"), 
        fg="#2a7e85", 
        bg="#f5f5f5",
        pady=15
    )
    title_label.pack(fill="x")

    # Combobox để chọn font
    font_frame = tk.Frame(root, bg="#f5f5f5")
    font_frame.pack(pady=5, fill="x", padx=30)
    
    font_var = tk.StringVar(root)
    font_combo = ttk.Combobox(
        font_frame, 
        textvariable=font_var, 
        values=[ "Montserrat", "Nunito", "OpenSans", "Pacifico"],
        state="readonly",
        width=30
    )
    font_combo.current(0)
    font_combo.pack(fill="x")

    # Khung cho file
    drop_frame = tk.Frame(
        root, 
        bg="#f0f0f0", 
        highlightbackground="#d0d0d0", 
        highlightthickness=1, 
        height=120
    )
    drop_frame.pack(pady=15, fill="x", padx=30)
    drop_frame.pack_propagate(False)  # Giữ kích thước cố định

    # Nút chọn file
    choose_btn = tk.Button(
        drop_frame, 
        text="Choose File...", 
        font=("OpenSans", 10),
        bg="#2a7e85", 
        fg="white",
        padx=15,
        pady=8,
        borderwidth=0,
        cursor="hand2"
    )
    choose_btn.pack(pady=(20, 10))
    choose_btn.config(command=select_file)

    # Hoặc kéo thả text (hiển thị nhưng không hoạt động chức năng kéo thả)
    drag_label = tk.Label(
        drop_frame, 
        text="Or drag and drop your PDF here", 
        fg="#777777", 
        bg="#f0f0f0",
        font=("OpenSans", 9)
    )
    drag_label.pack()

    # Hiển thị tên file đã chọn
    file_label = tk.Label(drop_frame, text="", bg="#f0f0f0", fg="#333333")
    file_label.pack(pady=5)

    # Nút Convert
    convert_btn = tk.Button(
        root, 
        text="Convert", 
        font=("OpenSans", 11, "bold"),
        bg="#2a7e85", 
        fg="white",
        padx=15,
        pady=10,
        borderwidth=0,
        cursor="hand2",
        width=25,
        command=convert_file
    )
    convert_btn.pack(pady=15)

    # Label hiển thị trạng thái
    status_label = tk.Label(root, text="", bg="#f5f5f5", fg="#555555")
    status_label.pack(pady=5)
    
    # Progress bar (hidden initially)
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    
    # Chạy giao diện
    root.mainloop()


if __name__ == "__main__":
    create_gui()