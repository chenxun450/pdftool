import os
import sys
import io
import time
import secrets
import string
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path


# PDF 处理库
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pikepdf
from pikepdf import Permissions, Encryption

# ---------- 工具函数 ----------
def find_chinese_font():
    """自动查找系统中可能的中文字体"""
    candidates = []
    if sys.platform.startswith('win'):
        candidates = [
            "C:/Windows/Fonts/simhei.ttf",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc"
        ]
    elif sys.platform.startswith('darwin'):  # macOS
        candidates = [
            "/System/Library/Fonts/PingFang.ttc",
            "/Library/Fonts/Arial Unicode.ttf",
            "/System/Library/Fonts/STHeiti Light.ttc"
        ]
    else:  # Linux
        candidates = [
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"
        ]

    for path in candidates:
        if os.path.exists(path):
            return path
    return None

def create_watermark(text, angle=270, font_size=16, font_path=None):
    """创建包含水印文字的PDF页面"""
    watermark_buffer = io.BytesIO()
    c = canvas.Canvas(watermark_buffer, pagesize=letter)

    # 注册字体
    try:
        if font_path and os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('CustomFont', font_path))
            c.setFont("CustomFont", font_size)
        else:
            # 尝试自动查找字体
            font_path = find_chinese_font()
            if font_path:
                pdfmetrics.registerFont(TTFont('AutoFont', font_path))
                c.setFont("AutoFont", font_size)
            else:
                # 无中文字体时，使用默认字体（会显示为乱码，但程序不崩溃）
                c.setFont("Helvetica", font_size)
    except Exception as e:
        print(f"字体加载失败：{e}")
        c.setFont("Helvetica", font_size)

    c.setFillColorRGB(0.4, 0.4, 0.4, 0.2)  # 灰色半透明
    width, height = letter

    # 水印布局：2列6行均匀分布
    cols, rows = 2, 6
    positions = [
        ((col + 0.5) / cols, (row + 0.5) / rows)
        for row in range(rows)
        for col in range(cols)
    ]

    for x_ratio, y_ratio in positions:
        draw_rotated_centred_text(c, text, width, height, angle,
                                   x_ratio=x_ratio, y_ratio=y_ratio)

    c.save()
    watermark_buffer.seek(0)
    return watermark_buffer

def draw_rotated_centred_text(c, text, width, height, angle, x_ratio=0.5, y_ratio=0.5):
    """在指定比例位置绘制旋转居中文字"""
    c.saveState()
    c.translate(width * x_ratio, height * y_ratio)
    c.rotate(angle)
    c.drawCentredString(0, 0, text)
    c.restoreState()

def add_text_watermark(input_pdf, output_pdf, watermark_text, angle=45, font_size=30, font_path=None):
    """给PDF文件添加水印"""
    watermark_pdf = create_watermark(watermark_text, angle, font_size, font_path)
    watermark_reader = PdfReader(watermark_pdf)
    watermark_page = watermark_reader.pages[0]

    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

def protect_pdf(input_path, output_path, owner_password=None):
    """
    设置PDF权限（禁止复制、编辑等），并加密。
    返回生成的所有者密码（如果未提供）。
    """
    with pikepdf.open(input_path) as pdf:
        permissions = Permissions(
            print_highres=False,
            print_lowres=False,
            extract=False,
            accessibility=False,
            modify_assembly=False,
            modify_other=False,
            modify_form=False,
            modify_annotation=False,
        )

        if not owner_password:
            alphabet = string.ascii_letters + string.digits
            owner_password = ''.join(secrets.choice(alphabet) for _ in range(16))

        pdf.save(
            output_path,
            encryption=Encryption(
                owner=owner_password,
                user="",
                allow=permissions
            )
        )
        return owner_password

# ---------- GUI 应用 ----------
class WatermarkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 水印与保护工具")
        self.root.geometry("700x550")

        # 变量
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.watermark_text = tk.StringVar()
        self.angle = tk.IntVar(value=45)
        self.font_size = tk.IntVar(value=16)
        self.font_path = tk.StringVar()
        self.owner_password = ""

        self.create_widgets()

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 输入文件
        ttk.Label(main_frame, text="输入 PDF 文件：").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_input).grid(row=0, column=2)

        # 输出文件
        ttk.Label(main_frame, text="输出 PDF 文件：").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="保存为...", command=self.browse_output).grid(row=1, column=2)

        # 水印文本
        ttk.Label(main_frame, text="水印文字（自定义部分）：").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.watermark_text, width=50).grid(row=2, column=1, padx=5)
        ttk.Label(main_frame, text="（最终水印格式：自考笔果题库{文字}产权保护传播必究）").grid(row=3, column=1, sticky=tk.W)

        # 角度
        ttk.Label(main_frame, text="旋转角度（度）：").grid(row=4, column=0, sticky=tk.W, pady=5)
        angle_frame = ttk.Frame(main_frame)
        angle_frame.grid(row=4, column=1, sticky=tk.W)
        ttk.Scale(angle_frame, from_=0, to=360, variable=self.angle, orient=tk.HORIZONTAL, length=200).pack(side=tk.LEFT)
        ttk.Label(angle_frame, textvariable=self.angle, width=5).pack(side=tk.LEFT, padx=5)

        # 字体大小
        ttk.Label(main_frame, text="字体大小：").grid(row=5, column=0, sticky=tk.W, pady=5)
        size_frame = ttk.Frame(main_frame)
        size_frame.grid(row=5, column=1, sticky=tk.W)
        ttk.Scale(size_frame, from_=10, to=80, variable=self.font_size, orient=tk.HORIZONTAL, length=200).pack(side=tk.LEFT)
        ttk.Label(size_frame, textvariable=self.font_size, width=5).pack(side=tk.LEFT, padx=5)

        # 字体文件
        ttk.Label(main_frame, text="字体文件（可选）：").grid(row=6, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.font_path, width=50).grid(row=6, column=1, padx=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_font).grid(row=6, column=2)

        # 开始按钮
        self.process_btn = ttk.Button(main_frame, text="开始处理", command=self.start_processing)
        self.process_btn.grid(row=7, column=1, pady=15)

        # 状态文本框
        self.status_text = tk.Text(main_frame, height=12, width=80, state=tk.DISABLED)
        self.status_text.grid(row=8, column=0, columnspan=3, pady=5)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=8, column=3, sticky=tk.NS)
        self.status_text['yscrollcommand'] = scrollbar.set

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.input_path.set(filename)
            # 自动建议输出文件名
            if not self.output_path.get():
                base = os.path.splitext(os.path.basename(filename))[0]
                self.output_path.set(os.path.join(os.path.dirname(filename), f"{base}_protected.pdf"))

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.output_path.set(filename)

    def browse_font(self):
        filename = filedialog.askopenfilename(filetypes=[("TrueType fonts", "*.ttf *.ttc")])
        if filename:
            self.font_path.set(filename)

    def log(self, message):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def start_processing(self):
        # 输入验证
        if not self.input_path.get() or not os.path.exists(self.input_path.get()):
            messagebox.showerror("错误", "请选择有效的输入PDF文件。")
            return
        if not self.output_path.get():
            messagebox.showerror("错误", "请指定输出文件路径。")
            return
        if not self.watermark_text.get():
            if not messagebox.askyesno("确认", "水印文字为空，将仅添加固定前缀和后缀。是否继续？"):
                return

        # 禁用按钮
        self.process_btn.config(state=tk.DISABLED, text="处理中...")
        self.log("开始处理...")

        # 在后台线程中运行，防止界面卡死
        thread = threading.Thread(target=self.process_pdf, daemon=True)
        thread.start()

    def process_pdf(self):
        try:
            input_file = self.input_path.get()
            output_file = self.output_path.get()
            custom_text = self.watermark_text.get()
            full_watermark = f"自考笔果题库{custom_text}产权保护传播必究"
            angle = self.angle.get()
            font_size = self.font_size.get()
            font_path = self.font_path.get() if self.font_path.get() else None

            # 1. 添加水印到临时文件
            temp_dir = os.path.join(os.path.dirname(__file__), "temp") if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), "temp")
            os.makedirs(temp_dir, exist_ok=True)
            temp_watermarked = os.path.join(temp_dir, f"temp_watermarked_{int(time.time())}.pdf")

            self.log("正在添加水印...")
            add_text_watermark(input_file, temp_watermarked, full_watermark, angle, font_size, font_path)

            # 2. 保护PDF
            self.log("正在设置权限与加密...")
            owner_pwd = protect_pdf(temp_watermarked, output_file)

            # 3. 清理临时文件
            try:
                os.remove(temp_watermarked)
            except:
                pass

            self.log(f"处理完成！\n输出文件：{output_file}\n所有者密码（用于解除限制）：{owner_pwd}\n请妥善保管该密码。")
            messagebox.showinfo("完成", f"PDF已成功处理。\n所有者密码：{owner_pwd}\n（已保存于日志中）")
        except Exception as e:
            self.log(f"错误：{e}")
            messagebox.showerror("错误", f"处理失败：{e}")
        finally:
            # 恢复按钮
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL, text="开始处理"))

if __name__ == "__main__":
    root = tk.Tk()
    app = WatermarkApp(root)
    root.mainloop()
