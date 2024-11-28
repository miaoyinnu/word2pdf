import os
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

class WordToPdfConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Word转PDF工具")
        self.window.geometry("600x400")
        
        self.selected_files = []
        self.setup_ui()
        
    def setup_ui(self):
        # 创建文件列表框
        self.files_frame = ttk.LabelFrame(self.window, text="已选择的文件")
        self.files_frame.pack(padx=10, pady=5, fill="both", expand=True)
        
        self.files_listbox = tk.Listbox(self.files_frame)
        self.files_listbox.pack(padx=5, pady=5, fill="both", expand=True)
        
        # 创建按钮
        self.buttons_frame = ttk.Frame(self.window)
        self.buttons_frame.pack(padx=10, pady=5, fill="x")
        
        ttk.Button(self.buttons_frame, text="选择文件", command=self.select_files).pack(side="left", padx=5)
        ttk.Button(self.buttons_frame, text="清除列表", command=self.clear_files).pack(side="left", padx=5)
        
        # 添加页码选项
        self.add_page_numbers = tk.BooleanVar()
        ttk.Checkbutton(self.buttons_frame, text="添加页码", variable=self.add_page_numbers).pack(side="left", padx=5)
        
        ttk.Button(self.buttons_frame, text="转换为PDF", command=self.convert_single).pack(side="left", padx=5)
        ttk.Button(self.buttons_frame, text="合并为PDF", command=self.merge_files).pack(side="left", padx=5)

    def select_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Word文件", "*.docx;*.doc")]
        )
        if files:
            if len(self.selected_files) + len(files) > 5:
                messagebox.showwarning("警告", "最多只能选择5个文件！")
                return
            
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))

    def clear_files(self):
        self.selected_files.clear()
        self.files_listbox.delete(0, tk.END)

    def convert_to_pdf(self, word_path, pdf_path):
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False  # 设置Word程序不可见
            
            # 确保文件路径是绝对路径
            word_path = os.path.abspath(word_path)
            pdf_path = os.path.abspath(pdf_path)
            
            doc = word.Documents.Open(word_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 表示 PDF 格式
            doc.Close()
            
        except Exception as e:
            if word:
                try:
                    word.Quit()
                except:
                    pass
            raise Exception(f"转换失败: {str(e)}")
            
        finally:
            if word:
                try:
                    word.Quit()
                except:
                    pass

    def add_page_numbers_to_pdf(self, input_pdf, output_pdf):
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.drawString(300, 30, str(page_num + 1))
            can.save()
            packet.seek(0)
            
            watermark = PdfReader(packet)
            page.merge_page(watermark.pages[0])
            writer.add_page(page)
            
        with open(output_pdf, 'wb') as output_file:
            writer.write(output_file)

    def convert_single(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择文件！")
            return
            
        if len(self.selected_files) > 1:
            messagebox.showwarning("警告", "单个转换只能选择一个文件！")
            return
            
        input_file = self.selected_files[0]
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")]
        )
        
        if output_file:
            try:
                self.convert_to_pdf(input_file, output_file)
                if self.add_page_numbers.get():
                    temp_file = output_file + ".temp"
                    os.rename(output_file, temp_file)
                    self.add_page_numbers_to_pdf(temp_file, output_file)
                    os.remove(temp_file)
                messagebox.showinfo("成功", "文件转换完成！")
            except Exception as e:
                messagebox.showerror("错误", f"转换失败：{str(e)}")

    def merge_files(self):
        if len(self.selected_files) < 2:
            messagebox.showwarning("警告", "请至少选择两个文件进行合并！")
            return
            
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")]
        )
        
        if output_file:
            try:
                # 创建临时目录
                temp_dir = "temp_pdfs"
                os.makedirs(temp_dir, exist_ok=True)
                
                # 转换所有Word文件为PDF
                pdf_files = []
                for word_file in self.selected_files:
                    pdf_file = os.path.join(temp_dir, os.path.basename(word_file) + ".pdf")
                    self.convert_to_pdf(word_file, pdf_file)
                    pdf_files.append(pdf_file)
                
                # 合并PDF文件
                merger = PdfMerger()
                for pdf_file in pdf_files:
                    merger.append(pdf_file)
                    
                merger.write(output_file)
                merger.close()
                
                # 如果需要添加页码
                if self.add_page_numbers.get():
                    temp_file = output_file + ".temp"
                    os.rename(output_file, temp_file)
                    self.add_page_numbers_to_pdf(temp_file, output_file)
                    os.remove(temp_file)
                
                # 清理临时文件
                for pdf_file in pdf_files:
                    os.remove(pdf_file)
                os.rmdir(temp_dir)
                
                messagebox.showinfo("成功", "文件合并完成！")
            except Exception as e:
                messagebox.showerror("错误", f"合并失败：{str(e)}")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = WordToPdfConverter()
    app.run() 