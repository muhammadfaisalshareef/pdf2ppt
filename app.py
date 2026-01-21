# -*- coding: utf-8 -*-
import webbrowser
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys
import logging
import os
import contextlib
from datetime import datetime

# 导入核心逻辑
try:
    import pdf2ppt
except ImportError:
    # 如果是在同一目录，直接导入；如果打包后路径变化，可能需要调整 path
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import pdf2ppt

# 设置外观模式
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

# ==================== 日志重定向工具 ====================
class TextHandler(logging.Handler):
    """
    自定义 Logging Handler，将日志输出到 GUI 的文本框
    """
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(ctk.END, msg + "\n")
            self.text_widget.see(ctk.END)
            self.text_widget.configure(state='disabled')
        # 确保在主线程更新 UI
        self.text_widget.after(0, append)

# ==================== GUI应用 ====================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF 转 PPT 工具 v0.3")
        self.geometry("700x650")

        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Top: Inputs
        self.grid_rowconfigure(1, weight=0)  # Middle: Button
        self.grid_rowconfigure(2, weight=1)  # Bottom: Log

        # ==================== 1. 上方输入区 ====================
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.input_frame.grid_columnconfigure(1, weight=1)

        # 1.1 MinerU Token
        self.label_token = ctk.CTkLabel(self.input_frame, text="MinerU Token:", anchor="w")
        self.label_token.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entry_token = ctk.CTkEntry(self.input_frame, placeholder_text="请输入您的 MinerU Token")
        self.entry_token.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.btn_token_url = ctk.CTkButton(self.input_frame, text="获取Token", width=80, command=self.open_token_url)
        self.btn_token_url.grid(row=0, column=2, padx=10, pady=10)

        # 1.2 PDF 文件选择
        self.label_pdf = ctk.CTkLabel(self.input_frame, text="PDF 文件:", anchor="w")
        self.label_pdf.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.entry_pdf = ctk.CTkEntry(self.input_frame, placeholder_text="请选择 PDF 文件")
        self.entry_pdf.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.btn_pdf = ctk.CTkButton(self.input_frame, text="浏览...", width=80, command=self.browse_pdf)
        self.btn_pdf.grid(row=1, column=2, padx=10, pady=10)

        # 1.3 PPT 输出路径
        self.label_ppt = ctk.CTkLabel(self.input_frame, text="PPT 输出:", anchor="w")
        self.label_ppt.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.entry_ppt = ctk.CTkEntry(self.input_frame, placeholder_text="请选择保存路径")
        self.entry_ppt.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.btn_ppt = ctk.CTkButton(self.input_frame, text="浏览...", width=80, command=self.browse_ppt_save)
        self.btn_ppt.grid(row=2, column=2, padx=10, pady=10)

        # 1.4 PPT 比例 和 去水印选项
        self.label_ratio = ctk.CTkLabel(self.input_frame, text="PPT 比例:", anchor="w")
        self.label_ratio.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        
        # 比例选择器
        self.ratio_var = ctk.StringVar(value="16:9")
        self.ratio_combobox = ctk.CTkComboBox(
            self.input_frame, 
            values=["16:9", "4:3"],
            variable=self.ratio_var,
            width=120
        )
        self.ratio_combobox.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        # 去水印复选框（在同一行右侧）
        self.remove_watermark_var = ctk.BooleanVar(value=True)
        self.remove_watermark_checkbox = ctk.CTkCheckBox(
            self.input_frame,
            text="去水印",
            variable=self.remove_watermark_var
        )
        self.remove_watermark_checkbox.grid(row=3, column=1, padx=(140, 10), pady=10, sticky="w")

        # ==================== 2. 中间按钮区 ====================
        self.btn_generate = ctk.CTkButton(
            self, 
            text="开始转换", 
            height=50, 
            font=("Microsoft YaHei UI", 16, "bold"),
            command=self.start_conversion_thread
        )
        self.btn_generate.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        # ==================== 3. 下方日志区 ====================
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(self.log_frame, state="disabled", font=("Consolas", 12))
        self.log_textbox.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        # 配置日志重定向
        self.setup_logging()

    def open_token_url(self):
        webbrowser.open("https://mineru.net/apiManage/token")

    def setup_logging(self):
        # 创建自定义 Handler
        text_handler = TextHandler(self.log_textbox)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        # 获取 pdf2ppt 的 logger (它使用 root logger 或者其它)
        # 由于 pdf2ppt.py 里用了 logging.info(...)，它是往 root logger 发的
        # 但我们在 pdf2ppt.main() 里配置了 basicConfig。
        # 当作为模块导入时，main() 不会运行，所以 basicConfig 没配。
        # 我们在这里配置 root logger。
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.addHandler(text_handler)
        
        # 同时保留控制台输出（可选）
        console_handler = logging.StreamHandler()
        logger.addHandler(console_handler)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.entry_pdf.delete(0, ctk.END)
            self.entry_pdf.insert(0, filename)
            
            # 自动建议 PPT 路径
            base_name = os.path.splitext(filename)[0]
            ppt_name = base_name + ".pptx"
            self.entry_ppt.delete(0, ctk.END)
            self.entry_ppt.insert(0, ppt_name)

    def browse_ppt_save(self):
        filename = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")])
        if filename:
            self.entry_ppt.delete(0, ctk.END)
            self.entry_ppt.insert(0, filename)

    def start_conversion_thread(self):
        # 禁用按钮
        self.btn_generate.configure(state="disabled", text="正在转换...")
        
        thread = threading.Thread(target=self.run_conversion)
        thread.daemon = True
        thread.start()

    def run_conversion(self):
        try:
            token = self.entry_token.get().strip()
            pdf_path = self.entry_pdf.get().strip()
            ppt_path = self.entry_ppt.get().strip()
            ratio_str = self.ratio_var.get()

            if not token:
                raise ValueError("请填写 MinerU Token")
            if not pdf_path:
                raise ValueError("请选择 PDF 文件")
            if not ppt_path:
                raise ValueError("请选择 PPT 输出路径")

            # 解析比例
            if ratio_str == "4:3":
                w, h = 10, 7.5  # 4:3 标准
            else:
                w, h = 16, 9    # 16:9 标准

            logging.info("--- 开始任务 ---")
            logging.info(f"输入: {pdf_path}")
            logging.info(f"输出: {ppt_path}")
            
            # 调用 pdf2ppt 逻辑
            # 注意：GUI 模式下强制不使用缓存（user requirement: GUI page unused cache option, default closed）
            try:
                # 尝试调用
                pdf2ppt.convert_pdf_to_ppt(
                    pdf_input_path=pdf_path,
                    ppt_output_path=ppt_path,
                    mineru_token=token,
                    ppt_slide_width=w,
                    ppt_slide_height=h,
                    use_cache=False, 
                    cache_dir="temp", # 仅用于调试
                    remove_watermark=self.remove_watermark_var.get()
                )
                messagebox.showinfo("成功", "PPT 转换完成！")
            except Exception as e:
                logging.error(f"转换出错: {e}")
                messagebox.showerror("错误", f"转换失败: {e}")

        except ValueError as ve:
            messagebox.showwarning("提示", str(ve))
        except Exception as e:
            messagebox.showerror("未知错误", str(e))
        finally:
            # 恢复按钮
            self.btn_generate.configure(state="normal", text="开始转换")
            logging.info("--- 任务结束 ---")

if __name__ == "__main__":
    app = App()
    app.mainloop()
