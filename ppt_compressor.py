import os
import sys
import zipfile
import io
import threading
import tkinter as as_tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw

# --- 配置 ---
COMPRESSION_PROFILES = {
    "🏆 智能高清 (推荐)":      {"max_width": 2048, "quality": 85},
    "⚖️ 均衡模式 (适合传阅)":   {"max_width": 1600, "quality": 75},
    "📉 强力压缩 (适合手机)":   {"max_width": 1280, "quality": 60},
    "🔥 极限压缩 (最小体积)":   {"max_width": 1024, "quality": 50}
}

# --- 自定义控件：高清抗锯齿加载器 ---
class SmoothSpinner(as_tk.Label):
    def __init__(self, parent, size=28, color="#0078D7", bg_color=None):
        # 如果未指定背景色，尝试获取父容器背景，失败则用白色
        if bg_color is None:
            try: bg_color = ttk.Style().lookup("TFrame", "background")
            except: bg_color = "#ffffff"
            
        super().__init__(parent, background=bg_color, borderwidth=0)
        self.size = size
        self.color = color
        self.bg_color = bg_color
        self.frames = []
        self.current_frame = 0
        self.is_spinning = False
        
        # 核心：预生成 30 帧高清旋转图片
        self._generate_frames()
        
        # 设置初始帧
        self.config(image=self.frames[0])

    def _generate_frames(self):
        """生成一圈高清平滑的旋转动画帧"""
        # 放大倍数 (Supersampling)，4倍采样可以完全消除锯齿
        scale = 4 
        real_size = self.size * scale
        line_width = 3 * scale
        
        # 每一帧旋转 12 度，共 30 帧
        for i in range(30):
            # 创建透明背景图
            # 注意：Tkinter 对透明支持有限，最好用背景色填充
            img = Image.new('RGBA', (real_size, real_size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # 计算圆的边界
            padding = 2 * scale
            bbox = (padding, padding, real_size - padding, real_size - padding)
            
            # 绘制圆弧 (缺口圆环)
            #start_angle = -90 - (i * 12)
            #draw.arc(bbox, start=start_angle, end=start_angle + 270, fill=self.color, width=line_width)
            
            # 绘制更高级的"追尾"效果 (Gradient Arc)
            # 我们画一段弧，并在末端画圆头
            start_angle = (i * 12) 
            end_angle = start_angle + 280
            
            # 画一个圆环
            draw.arc(bbox, start=start_angle, end=end_angle, fill=self.color, width=line_width)
            
            # 缩小回正常尺寸 (LANCZOS 滤镜是抗锯齿的关键)
            img = img.resize((self.size, self.size), Image.Resampling.LANCZOS)
            
            # 创建混合背景的图像 (解决 Tkinter 边缘黑边问题)
            bg_img = Image.new('RGB', (self.size, self.size), self.bg_color)
            bg_img.paste(img, (0, 0), mask=img)
            
            self.frames.append(ImageTk.PhotoImage(bg_img))

    def start(self):
        if not self.is_spinning:
            self.is_spinning = True
            self.animate()

    def stop(self):
        self.is_spinning = False

    def animate(self):
        if not self.is_spinning:
            return
            
        # 切换下一帧
        self.current_frame = (self.current_frame + 1) % len(self.frames)
        self.config(image=self.frames[self.current_frame])
        
        # 30ms 刷新一次，非常流畅
        self.after(30, self.animate)


# --- 主程序 ---
class PPTCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT 自由压缩工具 (极致画质版)")
        self.root.geometry("600x420")
        self.root.resizable(False, False)


        try:
            # 获取资源文件的绝对路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的 exe 运行
                base_path = sys._MEIPASS
            else:
                # 如果是 Python 脚本运行
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "app.ico")
            
            # 设置窗口图标
            self.root.iconbitmap(icon_path)
        except Exception as e:
            # 如果找不到图标，就不设置，避免程序闪退
            print(f"图标加载失败: {e}")
        
        # 变量
        self.file_path_var = ttk.StringVar()
        self.status_var = ttk.StringVar(value="准备就绪")
        self.progress_var = ttk.IntVar(value=0)
        
        self.is_running = False
        self.stop_event = threading.Event()
        
        self.create_widgets()

    def create_widgets(self):
        # 1. 标题区
        header = ttk.Frame(self.root, padding=20)
        header.pack(fill=X)
        ttk.Label(header, text="PPT 智能压缩工具", font=("微软雅黑", 18, "bold"), bootstyle="primary").pack(side=LEFT)
        ttk.Label(header, text="v7.0 Ultra", font=("Arial", 10), bootstyle="secondary").pack(side=LEFT, padx=10, pady=(10,0))

        # 2. 文件选择
        input_frame = ttk.Labelframe(self.root, text="第一步：选择文件", padding=15, bootstyle="info")
        input_frame.pack(fill=X, padx=20, pady=5)
        
        self.entry = ttk.Entry(input_frame, textvariable=self.file_path_var)
        self.entry.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))
        ttk.Button(input_frame, text="📂 浏览", bootstyle="outline-info", command=self.select_file).pack(side=LEFT)

        # 3. 压缩选项
        opt_frame = ttk.Labelframe(self.root, text="第二步：选择压缩强度", padding=15, bootstyle="warning")
        opt_frame.pack(fill=X, padx=20, pady=10)
        
        self.combo_mode = ttk.Combobox(opt_frame, values=list(COMPRESSION_PROFILES.keys()), state="readonly", bootstyle="warning")
        self.combo_mode.current(0)
        self.combo_mode.pack(fill=X)

        # 4. 进度条区域
        progress_frame = ttk.Frame(self.root, padding=(20, 0))
        progress_frame.pack(fill=X)
        
        self.progress = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            bootstyle="success-striped", 
            mode='determinate'
        )
        
        # 5. 底部状态与按钮
        action_frame = ttk.Frame(self.root, padding=20)
        action_frame.pack(fill=X, side=BOTTOM)
        
        # 左侧容器：包含 Spinner 和 文本
        status_container = ttk.Frame(action_frame)
        status_container.pack(side=LEFT)

        # 实例化高清 Spinner (这里会自动获取父容器背景色，实现无缝融合)
        # 获取当前主题的背景色，通常是 #ffffff 或 #f8f9fa
        bg_col = ttk.Style().lookup("TFrame", "background")
        self.spinner = SmoothSpinner(status_container, size=24, color="#0078D7", bg_color=bg_col)
        # 默认隐藏
        
        self.status_lbl = ttk.Label(status_container, textvariable=self.status_var, bootstyle="secondary", font=("Consolas", 9))
        self.status_lbl.pack(side=LEFT, padx=10)
        
        # 右侧按钮
        self.btn_run = ttk.Button(
            action_frame, 
            text="开始压缩 🚀", 
            bootstyle="primary", 
            width=20, 
            command=self.toggle_process
        )
        self.btn_run.pack(side=RIGHT)

    def select_file(self):
        if self.is_running: return 
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.file_path_var.set(path)
            self.status_var.set("已加载文件")
            self.progress_var.set(0)

    def toggle_process(self):
        if not self.is_running:
            self.start_thread()
        else:
            if messagebox.askyesno("确认", "确定要停止压缩吗？"):
                self.status_var.set("正在停止...")
                self.stop_event.set()
                self.btn_run.config(state=DISABLED)

    def update_ui_running(self, is_running):
        self.is_running = is_running
        if is_running:
            self.btn_run.config(text="⏹ 停止压缩", bootstyle="danger", state=NORMAL)
            self.entry.config(state=DISABLED)
            self.combo_mode.config(state=DISABLED)
            self.progress.pack(fill=X, pady=10)
            
            # 显示并启动 Spinner
            self.spinner.pack(side=LEFT)
            self.spinner.start()
        else:
            self.btn_run.config(text="开始压缩 🚀", bootstyle="primary", state=NORMAL)
            self.entry.config(state=NORMAL)
            self.combo_mode.config(state="readonly")
            self.stop_event.clear()
            
            # 停止并隐藏
            self.spinner.stop()
            self.spinner.pack_forget()

    def update_progress(self, current, total, filename):
        if total == 0: return
        percent = (current / total) * 100
        # 截断长文件名
        short_name = filename if len(filename) < 25 else "..." + filename[-25:]
        
        self.root.after(0, lambda: self.progress_var.set(percent))
        self.root.after(0, lambda: self.status_var.set(f"[{int(percent)}%] 处理中: {short_name}"))

    def start_thread(self):
        input_path = self.file_path_var.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("错误", "请先选择有效的 PPTX 文件")
            return
        
        mode_name = self.combo_mode.get()
        settings = COMPRESSION_PROFILES[mode_name]
        
        self.update_ui_running(True)
        self.status_var.set("正在分析文件结构...")
        
        t = threading.Thread(target=self.run_logic, args=(input_path, settings, self.update_progress))
        t.daemon = True
        t.start()

    def run_logic(self, input_path, settings, progress_callback):
        try:
            dir_name, file_name = os.path.split(input_path)
            base_name, ext = os.path.splitext(file_name)
            
            tag = "高清"
            if settings['quality'] < 70: tag = "强力"
            if settings['quality'] < 60: tag = "极限"
            output_path = os.path.join(dir_name, f"{base_name}_{tag}压缩{ext}")
            
            success, msg = self.compress_pptx_core(input_path, output_path, settings, progress_callback)
            self.root.after(0, lambda: self.on_finish(success, msg, input_path, output_path))
            
        except Exception as e:
            self.root.after(0, lambda: self.on_error(str(e)))

    def on_finish(self, success, msg, input_path, output_path):
        self.update_ui_running(False)
        if success:
            try:
                org_size = os.path.getsize(input_path) / 1024 / 1024
                new_size = os.path.getsize(output_path) / 1024 / 1024
                reduction = (1 - new_size / org_size) * 100
                self.status_var.set(f"完成！体积减少 {reduction:.1f}%")
                self.progress_var.set(100)
                messagebox.showinfo("搞定", f"✅ 压缩成功！\n体积减少: {reduction:.1f}%")
            except:
                messagebox.showinfo("搞定", "压缩成功！")
        else:
            if msg == "CANCELLED":
                self.status_var.set("已取消操作")
                self.progress_var.set(0)
                if os.path.exists(output_path):
                    try: os.remove(output_path)
                    except: pass
            else:
                self.status_var.set("失败")
                messagebox.showerror("错误", f"压缩失败: {msg}")

    def on_error(self, error_msg):
        self.update_ui_running(False)
        self.status_var.set("发生错误")
        messagebox.showerror("错误", error_msg)

# --- 核心压缩逻辑 ---
    def compress_image(self, image_data, filename, max_width, quality):
        try:
            img = Image.open(io.BytesIO(image_data))
            img_format = img.format
            
            # 过滤掉非图片或无法处理的格式
            if img_format not in ['JPEG', 'PNG', 'TIFF', 'BMP', 'GIF']: 
                return image_data

            # 1. 尺寸调整 (Resizing)
            width, height = img.size
            if width > max_width:
                ratio = max_width / width
                new_height = int(height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
            output_buffer = io.BytesIO()

            # 2. 针对不同格式的压缩策略
            if img_format == 'PNG':
                # 【关键修改】针对 PNG 进行色彩量化 (Quantize)
                # 如果是强力或极限模式 (quality < 70)，将 PNG 转为 256 色索引图
                # 这能保留透明度，同时体积减少 70% 以上
                if quality < 75:
                    # 必须先转为 RGBA 确保透明度被处理
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    
                    # quantize 需要 fast octree 算法，dither=Image.Dither.FLOYDSTEINBERG 能防色带
                    # colors=256 是 8-bit，colors=128 会更小
                    img = img.quantize(colors=256, method=2, dither=1)
                    
                    # 保存为 PNG
                    img.save(output_buffer, format='PNG', optimize=True)
                else:
                    # 高画质模式下，仅做 resize 和基础优化
                    img.save(output_buffer, format='PNG', optimize=True)

            else:
                # JPEG/其他格式的处理
                if img.mode != 'RGB': 
                    img = img.convert('RGB')
                
                # 处理 EXIF 旋转问题 (可选，防止压缩后图片倒转，这里暂略)
                img.save(output_buffer, format='JPEG', quality=quality, optimize=True)

            # 3. 防反向压缩检查
            # 如果压缩后反而比原图大（极少见），则返回原图
            compressed_data = output_buffer.getvalue()
            if len(compressed_data) >= len(image_data):
                return image_data
                
            return compressed_data
            
        except Exception as e:
            # print(f"Error compressing {filename}: {e}") # 调试用
            return image_data

    def compress_pptx_core(self, input_path, output_path, settings, callback=None):
        try:
            quality = settings['quality']
            max_width = settings['max_width']
            with zipfile.ZipFile(input_path, 'r') as zin:
                file_list = zin.infolist()
                total_files = len(file_list)
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for index, item in enumerate(file_list):
                        if self.stop_event.is_set(): return False, "CANCELLED"
                        
                        file_content = zin.read(item.filename)
                        if item.filename.startswith('ppt/media/') and \
                           item.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp')):
                            
                            compressed_content = self.compress_image(file_content, item.filename, max_width, quality)
                            zout.writestr(item, compressed_content)
                        else:
                            zout.writestr(item, file_content)
                        
                        if callback: callback(index + 1, total_files, item.filename)
            return True, "SUCCESS"
        except Exception as e:
            return False, str(e)

if __name__ == "__main__":
    app = ttk.Window(themename="cosmo") 
    PPTCompressorApp(app)
    app.mainloop()