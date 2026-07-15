#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox
from io import BytesIO
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import queue
from collections import deque
from html.parser import HTMLParser

# ==========================================
# 1. 插件路径与第三方依赖加载 (自适应策略)
# ==========================================
_PLUGIN_DIR = Path(__file__).resolve().parent
_VENDOR_DIR = _PLUGIN_DIR / "vendor"

def setup_environment():
    if not _VENDOR_DIR.exists():
        _VENDOR_DIR.mkdir(parents=True, exist_ok=True)
        
    vendor_path = str(_VENDOR_DIR)
    if vendor_path not in sys.path:
        sys.path.insert(0, vendor_path)
            
    if sys.platform == 'win32' and hasattr(os, 'add_dll_directory'):
        try:
            os.add_dll_directory(vendor_path)
            for item in _VENDOR_DIR.iterdir():
                # 兼容包含动态链接库的包 (如 .libs 后缀或 imagequant 扩展)
                if item.is_dir() and (item.name.endswith('.libs') or item.name == 'imagequant'):
                    os.add_dll_directory(str(item))
        except Exception:
            pass

setup_environment()

# ==========================================
# 2. 依赖检查与手动安装提示
# ==========================================
def check_dependencies():
    try:
        import yaml
        import PIL
        from PIL import Image
        from resvg_py import svg_to_bytes
        import cffi
        import imagequant
        
        # 版本检查逻辑
        def _parse_version(v_str):
            match = re.search(r'^(\d+\.\d+(\.\d+)?)', str(v_str))
            return tuple(map(int, match.group(1).split('.'))) if match else (0, 0, 0)
            
        if hasattr(yaml, '__version__') and _parse_version(yaml.__version__) < (5, 1):
            return f"pyyaml 版本过低 (当前 {yaml.__version__}，需 >= 5.1)"
            
        if hasattr(PIL, '__version__') and _parse_version(PIL.__version__) < (8, 0):
            return f"Pillow 版本过低 (当前 {PIL.__version__}，需 >= 8.0.0)"
            
        return None
    except Exception as e:
        return str(e)


class ConfigError(RuntimeError): pass

def svg2img(svg_path: Path, svg_width: int):
    from resvg_py import svg_to_bytes
    from PIL import Image, UnidentifiedImageError
    with BytesIO(svg_to_bytes(svg_path=str(svg_path), width=svg_width)) as svg_buffer:
        try:
            with Image.open(svg_buffer) as img:
                return img.convert("RGBA").copy()
        except Exception as e:
            raise UnidentifiedImageError(f"无法加载水印图片：{e}")

# --- 3. 核心水印处理引擎 ---
class SigilWatermarker:
    default = {
        "threads": 4,
        "filter": {
            "exclude_html": ["cover.xhtml", "nav.xhtml"],
            "exclude_images": ["logo.png"],
            "exclude_classes": ["no-watermark"],
            "exclude_tags": ["p"]
        },
        "process": {
            "target_formats": ["jpg", "jpeg", "png", "webp"], 
            "output_format": "webp", 
            "quality": 95,
        },
        "watermark": {
            "image_filename": "watermark.png",
            "width": {"unit_is_px_or_percent": "percent", "width_value": 10},
            "x_margin": {"begin_from_left_or_right": "right", "unit_is_px_or_percent": "percent", "margin_value": 5},
            "y_margin": {"begin_from_top_or_bottom": "bottom", "unit_is_px_or_percent": "percent", "margin_value": 5},
            "opacity": 0.5,
            "rotation": 0,
        },
    }

    def __init__(self, config_yaml_path: Path, max_epub_img_width: int = 1000) -> None:
        import yaml
        from PIL import Image

        if not config_yaml_path.exists():
            try:
                with open(config_yaml_path, "w", encoding="utf-8") as f:
                    yaml.dump(self.default, f, allow_unicode=True, sort_keys=False)
            except Exception as e:
                raise PermissionError(f"无法写入配置文件：{e}")
            raise ConfigError("配置文件不存在。\n已生成包含【黑名单规则】的默认配置。\n请检查 filter 节点后重新运行。")

        with open(config_yaml_path, "r", encoding="utf-8") as f:
            data = yaml.load(f, Loader=yaml.FullLoader)

        self.threads = data.get("threads", 4)
        filter_data = data.get("filter", {})
        
        self.exclude_html_stems = {Path(n).stem.lower() for n in (filter_data.get("exclude_html") or [])}
        self.exclude_images = {n.lower() for n in (filter_data.get("exclude_images") or [])}
        self.exclude_classes = set(filter_data.get("exclude_classes") or [])
        self.exclude_tags = {n.lower() for n in (filter_data.get("exclude_tags") or [])}

        process_data = data.get("process", {})
        self.target_formats = {fmt.lower() for fmt in process_data.get("target_formats", ["jpg", "jpeg", "png"])}
        self.output_format = process_data.get("output_format", "jpeg").lower()
        self.output_quality = int(process_data.get("quality", 95))

        wm_data = data.get("watermark", {})
        self.watermark_path = _PLUGIN_DIR / wm_data.get("image_filename", "watermark.svg")
        if not self.watermark_path.is_file():
            raise FileNotFoundError("找不到水印图片！请确保路径正确。")

        self.watermark_opacity = float(wm_data.get("opacity", 0.5))
        self.watermark_rotation = float(wm_data.get("rotation", 0))

        wm_width_data = wm_data.get("width", {})
        self.watermark_width = int(wm_width_data.get("width_value", 10))
        self.watermark_width_is_percent = (wm_width_data.get("unit_is_px_or_percent") != "px")

        wm_x_data = wm_data.get("x_margin", {})
        self.watermark_x_begin = wm_x_data.get("begin_from_left_or_right", "right")
        self.watermark_x_margin = int(wm_x_data.get("margin_value", 5))
        self.watermark_x_margin_is_percent = (wm_x_data.get("unit_is_px_or_percent") != "px")

        wm_y_data = wm_data.get("y_margin", {})
        self.watermark_y_begin = wm_y_data.get("begin_from_top_or_bottom", "bottom")
        self.watermark_y_margin = int(wm_y_data.get("margin_value", 5))
        self.watermark_y_margin_is_percent = (wm_y_data.get("unit_is_px_or_percent") != "px")

        if self.watermark_path.suffix.lower() == ".svg":
            self._watermark_buffer = svg2img(self.watermark_path, max(max_epub_img_width, 1000))
        else:
            with Image.open(self.watermark_path) as img:
                self._watermark_buffer = img.convert("RGBA").copy()

        # 初始化时预计算透明度与旋转
        self._prepare_watermark()

    def _prepare_watermark(self):
        from PIL import Image
        # 优化：重采样滤波器缓存 (兼容 PIL 8.0)
        self.resample_filter = getattr(Image.Resampling, 'LANCZOS', getattr(Image, 'LANCZOS', 1))
        
        wm = self._watermark_buffer.copy()
        
        # 优化：透明度预计算
        if self.watermark_opacity < 1:
            alpha = wm.getchannel('A')
            alpha = alpha.point(lambda p: int(p * self.watermark_opacity))
            wm.putalpha(alpha)
            
        # 优化：旋转预计算
        if self.watermark_rotation != 0:
            wm = wm.rotate(self.watermark_rotation, expand=True, resample=self.resample_filter)
            
        # 优化：绝对尺寸水印预缩放 (百分比水印在处理时动态缩放)
        if not self.watermark_width_is_percent:
            tw = self.watermark_width
            th = tw * wm.height // wm.width
            wm = wm.resize((tw, th), resample=self.resample_filter)
            
        self._prepared_watermark = wm
        self._watermark_buffer = None  # 释放原始大图引用，降低内存占用
        self._wm_cache = {}  # 引入按图片尺寸的水印缩放缓存字典

    def process_image_bytes(self, image_data: bytes) -> bytes:
        from PIL import Image
        import imagequant

        with Image.open(BytesIO(image_data)).convert("RGBA") as img:
            iw, ih = img.size
            
            # 使用缓存或预加载好的水印本体
            cache_key = (iw, ih)
            
            # 只有百分比模式才需要动态计算重采样
            if self.watermark_width_is_percent:
                if cache_key in self._wm_cache:
                    wm_p = self._wm_cache[cache_key]
                else:
                    wm_p = self._prepared_watermark
                    tw = iw * self.watermark_width // 100
                    th = tw * wm_p.height // wm_p.width
                    wm_p = wm_p.resize((tw, th), resample=self.resample_filter)
                    self._wm_cache[cache_key] = wm_p
            else:
                wm_p = self._prepared_watermark

            mx = iw * self.watermark_x_margin // 100 if self.watermark_x_margin_is_percent else self.watermark_x_margin
            my = ih * self.watermark_y_margin // 100 if self.watermark_y_margin_is_percent else self.watermark_y_margin
            px = mx if self.watermark_x_begin == "left" else iw - wm_p.width - mx
            py = my if self.watermark_y_begin == "top" else ih - wm_p.height - my

            img.paste(wm_p, (px, py), mask=wm_p)

            out_io = BytesIO()
            save_fmt = "JPEG" if self.output_format in ['jpg', 'jpeg'] else self.output_format.upper()
            if save_fmt == "JPEG": 
                img = img.convert("RGB")
            
            if save_fmt == "WEBP":
                # 微调 WebP 编码速度参数，去掉 optimize=True，增加 method=4
                img.save(out_io, format=save_fmt, quality=self.output_quality, method=4)
            elif self.output_quality < 100 and save_fmt == "PNG":
                imagequant.quantize_pil_image(img, max_quality=self.output_quality).save(out_io, format="PNG")
            else:
                img.save(out_io, format=save_fmt, quality=self.output_quality, optimize=True)
            
            return out_io.getvalue()

# --- 4. Python 原生 HTML 解析器引擎 (优化计数器机制) ---
class NativeEpubHTMLParser(HTMLParser):
    def __init__(self, wm_config, img_map):
        super().__init__(convert_charrefs=True)
        self.exclude_classes = wm_config.exclude_classes
        self.exclude_tags = wm_config.exclude_tags
        self.exclude_images = wm_config.exclude_images
        self.image_map = img_map
        
        self.extracted_ids = set()
        self.skipped_ids = set()
        
        # 优化：采用更高效的计数器代替高开销的集合操作
        self.tag_counter = {}
        self.class_counter = {}
        self.tag_stack = []

    def _increment(self, counter, key):
        counter[key] = counter.get(key, 0) + 1

    def _decrement(self, counter, key):
        if key in counter:
            counter[key] -= 1
            if counter[key] <= 0:
                del counter[key]

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        self._increment(self.tag_counter, tag)
        
        classes = []
        for k, v in attrs:
            if k == 'class':
                classes = v.split()
                for c in classes:
                    self._increment(self.class_counter, c)
                    
        self.tag_stack.append((tag, classes))
        self._check_image(tag, attrs)

    def handle_endtag(self, tag):
        tag = tag.lower()
        # 向上回溯弹出闭合标签对应的计数
        for i in range(len(self.tag_stack)-1, -1, -1):
            if self.tag_stack[i][0] == tag:
                popped = self.tag_stack[i:]
                del self.tag_stack[i:]
                for pt, pclasses in popped:
                    self._decrement(self.tag_counter, pt)
                    for pc in pclasses:
                        self._decrement(self.class_counter, pc)
                break

    def handle_startendtag(self, tag, attrs):
        tag = tag.lower()
        
        self._increment(self.tag_counter, tag)
        classes = []
        for k, v in attrs:
            if k == 'class':
                classes = v.split()
                for c in classes:
                    self._increment(self.class_counter, c)
            
        self._check_image(tag, attrs)
        
        self._decrement(self.tag_counter, tag)
        for c in classes:
            self._decrement(self.class_counter, c)

    def _check_image(self, tag, attrs):
        if tag not in ('img', 'image'):
            return
            
        src = None
        for k, v in attrs:
            if k in ('src', 'xlink:href', 'href'):
                src = v
                break

        if not src: 
            return
        
        img_basename = src.split('/')[-1].split('#')[0].split('?')[0].lower()
        if img_basename not in self.image_map: 
            return
        
        img_id = self.image_map[img_basename]
        
        # 优化：采用 dict_keys 视图做直接的 O(1) 交集匹配 (isdisjoint 运算)
        if (not self.exclude_classes.isdisjoint(self.class_counter.keys()) or 
            not self.exclude_tags.isdisjoint(self.tag_counter.keys()) or 
            img_basename in self.exclude_images):
            self.skipped_ids.add(img_id)
        else:
            self.extracted_ids.add(img_id)

# --- 5. 彻底重构的 UI 交互与异步流调度引擎 ---
class WatermarkApp:
    def __init__(self, root, bk, config, config_path):
        self.root = root
        self.bk = bk
        self.wm = config
        self.config_path = config_path
        
        self.phase = 1
        self.all_image_map = {}
        self.pre_selected_html_ids = set()
        
        self._init_data()
        self._build_ui()

    def _init_data(self):
        for img_info in self.bk.image_iter():
            img_id, href = img_info[0], img_info[1]
            basename = href.split('/')[-1].split('#')[0].split('?')[0].lower()
            self.all_image_map[basename] = img_id

        self.text_iter = list(self.bk.text_iter())
        try:
            for id_type, Id in self.bk.selected_iter():
                if id_type == "text": 
                    self.pre_selected_html_ids.add(Id)
        except Exception:
            pass

    def _build_ui(self):
        self.root.title("Watermarker-Air V1.1.5")
        self.root.geometry("600x590")
        self.root.minsize(580, 560)
        self.root.eval('tk::PlaceWindow . center')

        # -----------------------------
        # UI 样式全局定义区
        # -----------------------------
        style = ttk.Style()
        # 根据系统自适应最佳字体
        os_font = "微软雅黑" if sys.platform == "win32" else "Helvetica Neue" if sys.platform == "darwin" else "sans-serif"
        code_font = "Consolas" if sys.platform == "win32" else "Menlo" if sys.platform == "darwin" else "monospace"
        
        style.configure(".", font=(os_font, 10))
        style.configure("Treeview", rowheight=28, font=(code_font, 10))
        style.configure("Treeview.Heading", font=(os_font, 10, "bold"))
        
        # 定制按钮样式
        style.configure("Primary.TButton", font=(os_font, 10, "bold"), padding=(15, 6))
        style.configure("Secondary.TButton", font=(os_font, 10), padding=(10, 4))
        style.configure("TCheckbutton", font=(os_font, 10))

        # 主容器：所有元素被包裹在一个带内边距的主 Frame 中
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # -----------------------------
        # 1. 顶部说明区
        # -----------------------------
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.lbl_title = ttk.Label(header_frame, text="📚 选择需要扫描的 HTML 章节", font=(os_font, 13, "bold"))
        self.lbl_title.pack(anchor=tk.W, pady=(0, 4))
        
        self.lbl_hint = ttk.Label(
            header_frame, 
            text="按住 Ctrl 或 Shift 选择多个章节", 
            foreground="#666666"
        )
        self.lbl_hint.pack(anchor=tk.W)

        # -----------------------------
        # 2. 核心列表区
        # -----------------------------
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 改用 grid 布局来解决列表右侧滑动条导致的上下不对齐问题
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(list_frame, columns=("Data1", "Data2"), show="headings", selectmode="extended")
        self.tree.column("#0", width=0, stretch=tk.NO)
        # 设置宽度相同并全部开启拉伸，使其各占一半
        self.tree.column("Data1", anchor=tk.W, width=300, minwidth=150, stretch=tk.YES)
        self.tree.column("Data2", anchor=tk.W, width=300, minwidth=150, stretch=tk.YES)
        self.tree.heading("Data1", text="节点标识 (ID)")
        self.tree.heading("Data2", text="文件路径 (Href)")
        
        scroll_bar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_bar.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        scroll_bar.grid(row=0, column=1, sticky='ns')

        for index, (data1, data2) in enumerate(self.text_iter):
            item_id = self.tree.insert(parent="", index=tk.END, iid=f"html_{index}", values=(data1, data2))
            self.tree.selection_add(item_id) # 默认全选
            
        self.root.bind('<Control-a>', lambda e: self.tree.selection_add(self.tree.get_children()))
        self.root.bind('<Command-a>', lambda e: self.tree.selection_add(self.tree.get_children()))
        self.root.bind('<Return>', lambda e: self.run_action())
        
        # -----------------------------
        # 3. 操作按钮区 (重新布局：左右分栏)
        # -----------------------------
        self.action_frame = ttk.Frame(main_frame)
        self.action_frame.pack(fill=tk.X, pady=(10, 5))
        
        # 左侧区域：辅助设置
        left_controls = ttk.Frame(self.action_frame)
        left_controls.pack(side=tk.LEFT, fill=tk.Y)

        self.btn_open_config = ttk.Button(left_controls, text="⚙️ 唤起配置单", style="Secondary.TButton", command=self.open_config_file)
        self.btn_open_config.pack(side=tk.LEFT)
        
        # 右侧区域：主动作与二次确认
        right_controls = ttk.Frame(self.action_frame)
        right_controls.pack(side=tk.RIGHT, fill=tk.Y)

        self.btn_main = ttk.Button(right_controls, text="🚀 扫描并处理", style="Primary.TButton", command=self.run_action)
        self.btn_main.pack(side=tk.RIGHT)
        
        self.secondary_confirm_var = tk.BooleanVar(value=False)
        self.chk_confirm = ttk.Checkbutton(right_controls, text="🔍 图片二次确认", style="TCheckbutton", variable=self.secondary_confirm_var)
        self.chk_confirm.pack(side=tk.RIGHT, padx=(0, 15))

        # -----------------------------
        # 4. 进度条状态区
        # -----------------------------
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(10, 0))

        self.progress = ttk.Progressbar(status_frame, mode='determinate')
        self.progress.pack(side=tk.TOP, fill=tk.X)


    def open_config_file(self):
        import subprocess
        if not self.config_path.exists():
            messagebox.showwarning("提示", "配置文件不存在，请确保已初始化环境。")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(self.config_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', str(self.config_path)])
            else:
                subprocess.call(['xdg-open', str(self.config_path)])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开配置文件：{e}")

    def _set_ui_state(self, state):
        if state == tk.DISABLED:
            btn_text = "⏳ 正在处理中"
        else:
            btn_text = "✅ 确认并执行" if self.phase == 2 else "🚀 扫描并处理"
            
        self.btn_main.config(state=state, text=btn_text)
        self.tree.config(selectmode='none' if state == tk.DISABLED else 'extended')
        
        chk_state = tk.NORMAL if (state == tk.NORMAL and self.phase == 1) else tk.DISABLED
        self.chk_confirm.config(state=chk_state)
        self.btn_open_config.config(state=state)
        
        self.root.update_idletasks()

    def _worker_task(self, uid, img_data):
        """线程池内的安全任务封装，负责向 Queue 抛回结果"""
        try:
            result = self.wm.process_image_bytes(img_data)
            self.result_queue.put((uid, True, result))
        except Exception as e:
            self.result_queue.put((uid, False, e))

    def _fill_executor(self):
        """流式分块读取：保持活动任务数量不超过线程数的 2 倍，防止电子书过大撑爆内存"""
        target_active = self.wm.threads * 2
        while len(self.active_futures) < target_active and self.final_process_queue:
            uid = self.final_process_queue.popleft()
            try:
                # 必须在主线程触发宿主的 readfile
                img_data = self.bk.readfile(uid)
                future = self.executor.submit(self._worker_task, uid, img_data)
                self.active_futures[future] = uid
            except Exception as e:
                print(f"读取图片失败 {uid}: {e}")
                self.skipped_image_ids.add(uid)
                self._update_progress()

    def _update_progress(self):
        self.progress['value'] += 1
        # 优化：批量更新机制，限制重绘频率以降低 CPU 开销
        current_time = time.time()
        if current_time - self.last_update_time > 0.1:
            self.root.update_idletasks()
            self.last_update_time = current_time

    def _check_queue_and_refill(self):
        """after 定时轮询收集 Queue，实现完全非阻塞 UI 流畅过渡"""
        try:
            while True:
                uid, success, data_or_err = self.result_queue.get_nowait()
                
                if success:
                    try:
                        self.bk.writefile(uid, data_or_err)
                        self.success_ids.append(uid)
                    except Exception as e:
                        print(f"写入图片失败 {uid}: {e}")
                        self.skipped_image_ids.add(uid)
                else:
                    print(f"渲染失败 {uid}: {data_or_err}")
                    self.skipped_image_ids.add(uid)
                    
                self._update_progress()
        except queue.Empty:
            pass

        # 剥离已完成的任务池引用
        done_futures = [f for f in self.active_futures if f.done()]
        for f in done_futures:
            del self.active_futures[f]

        # 及时补充分块读取的任务
        self._fill_executor()

        # 全部结束判断
        if not self.active_futures and not self.final_process_queue:
            self.executor.shutdown(wait=False)
            self.root.update_idletasks()  # 确保最后一次进度正确触达 100%
            self._print_report(self.success_ids, self.skipped_image_ids)
            messagebox.showinfo("处理完成", f"批量水印添加完毕！\n\n✅ 成功打上水印: {len(self.success_ids)} 张\n🛑 跳过或被拦截: {len(self.skipped_image_ids)} 张\n\n详细名单请查看 Sigil 控制台输出。")
            self.root.destroy()
        else:
            self.root.after(50, self._check_queue_and_refill)

    def process_images_batch(self, final_process_queue, skipped_image_ids):
        if not final_process_queue:
            messagebox.showinfo("完成", "当前节点下未找到符合条件的图片，或图片已全部被黑名单拦截。")
            self.root.destroy()
            return

        self._set_ui_state(tk.DISABLED)
        self.progress['maximum'] = len(final_process_queue)
        self.progress['value'] = 0
        
        self.success_ids = []
        self.skipped_image_ids = skipped_image_ids
        # 优化：引入 collections.deque 替代 List 实现 O(1) 的 popleft()
        self.final_process_queue = deque(final_process_queue)
        self.last_update_time = time.time()
        
        self.result_queue = queue.Queue()
        self.active_futures = {}
        
        # 优化引擎：引入并发池，搭配流式读取与非阻塞 Queue。
        # 注意：Sigil 插件内置环境对于 ProcessPoolExecutor 进行 Pickle 极易引发模块重载错误与宿主闪退。
        # 这里统一降级绑定为 ThreadPoolExecutor（PIL 在底层处理时会自动释放 GIL 锁，完美模拟了多核多进程能力）。
        self.executor = ThreadPoolExecutor(max_workers=self.wm.threads)
        
        # 首次注水
        self._fill_executor()
        
        # 激活非阻塞 UI 轮询
        self.root.after(100, self._check_queue_and_refill)

    def _print_report(self, success_ids, skipped_image_ids):
        print("=" * 50)
        print("【水印处理报告】")
        print("=" * 50)
        print(f"\n✅ 处理成功的文件 ({len(success_ids)}个):")
        if success_ids:
            for sid in success_ids: print(f"  - {self.bk.id_to_href(sid).split('/')[-1]}")
        else: print("  (无)")

        print(f"\n🛑 拦截/略过的文件 ({len(skipped_image_ids)}个):")
        if skipped_image_ids:
            for xid in skipped_image_ids: print(f"  - {self.bk.id_to_href(xid).split('/')[-1]}")
        else: print("  (无)")
        print("=" * 50)

    def run_action(self):
        if self.phase == 1:
            selected_items = [self.tree.item(i)['values'] for i in self.tree.selection()]
            if not selected_items:
                messagebox.showwarning("提示", "请先在列表中选择需要扫描配图的章节！")
                return

            self._set_ui_state(tk.DISABLED)
            self.root.update_idletasks()
            
            selected_html_ids = [item[0] for item in selected_items]
            target_image_ids = set()
            skipped_image_ids = set()

            for html_id in selected_html_ids:
                html_data = self.bk.readfile(html_id)
                parser = NativeEpubHTMLParser(self.wm, self.all_image_map)
                parser.feed(html_data)
                
                href = self.bk.id_to_href(html_id)
                html_stem = Path(href.split('/')[-1].split('#')[0].split('?')[0]).stem.lower()
                
                if html_stem in self.wm.exclude_html_stems:
                    skipped_image_ids.update(parser.extracted_ids)
                    skipped_image_ids.update(parser.skipped_ids)
                else:
                    target_image_ids.update(parser.extracted_ids)
                    skipped_image_ids.update(parser.skipped_ids)

            target_image_ids.difference_update(skipped_image_ids)

            final_process_queue = []
            for img_id in target_image_ids.copy():
                ext = self.bk.id_to_href(img_id).rsplit('.')[-1].lower()
                if ext in self.wm.target_formats:
                    final_process_queue.append(img_id)
                else:
                    skipped_image_ids.add(img_id)

            if self.secondary_confirm_var.get():
                self.phase = 2
                self.root.title("Watermarker-Air - 目标确认阶段")
                self.lbl_title.config(text="🎯 核对并确认待处理的图片清单")
                self.lbl_hint.config(
                    text="下方展示了将被打上水印的所有图片（已被黑名单过滤）", 
                    foreground="#666666"
                )
                
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                for index, img_info in enumerate(self.bk.image_iter()):
                    img_id, href = img_info[0], img_info[1]
                    href_display = href if href else ""
                    item_id = self.tree.insert(parent="", index=tk.END, iid=f"img_{index}", values=(img_id, href_display))
                    if img_id in final_process_queue:
                        self.tree.selection_add(item_id)
                
                self._set_ui_state(tk.NORMAL) 
            else:
                self.process_images_batch(final_process_queue, skipped_image_ids)

        elif self.phase == 2:
            selected_items = [self.tree.item(i)['values'] for i in self.tree.selection()]
            if not selected_items:
                messagebox.showwarning("提示", "未选择任何需要处理的图片，请重新选择或关闭窗口！")
                return
                
            final_process_queue = []
            skipped_image_ids = set()
            
            for item in selected_items:
                img_id = str(item[0])
                href = str(item[1]) if item[1] else ""
                
                if href:
                    ext = href.rsplit('.')[-1].lower()
                    if ext in self.wm.target_formats:
                        final_process_queue.append(img_id)
                    else:
                        skipped_image_ids.add(img_id)
                else:
                    final_process_queue.append(img_id)
                    
            self.process_images_batch(final_process_queue, skipped_image_ids)

# --- 6. Sigil 插件入口 ---
def run(bk):
    # 增强跨平台 UI 兼容性：主动设置 Windows 高 DPI 缩放感知
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    err_msg = check_dependencies()
    
    if err_msg:
        # 将错误弹窗逻辑嵌入到入口处，防止阻塞加载
        err_root = tk.Tk()
        err_root.withdraw()
        
        guide_win = tk.Toplevel()
        guide_win.title("⚠️ 插件环境异常")
        guide_win.geometry("600x360")
        
        # 跨平台字体回退机制
        ui_font = ("微软雅黑" if sys.platform == "win32" else "Helvetica Neue" if sys.platform == "darwin" else "sans-serif", 10)
        code_font = ("Consolas" if sys.platform == "win32" else "Menlo" if sys.platform == "darwin" else "monospace", 10)
        
        # 窗口居中
        guide_win.update_idletasks()
        win_x = (guide_win.winfo_screenwidth() // 2) - (600 // 2)
        win_y = (guide_win.winfo_screenheight() // 2) - (320 // 2)
        guide_win.geometry(f'+{win_x}+{win_y}')
        
        top_frame = tk.Frame(guide_win, padx=25, pady=20)
        top_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(top_frame, text="由于缺失必要的核心组件，插件无法启动。", font=(ui_font[0], 12, "bold"), fg="#d9534f").pack(anchor=tk.W, pady=(0, 10))
        
        info_text = (
            f"具体拦截原因: {err_msg}\n\n"
            f"请按下 Win+R 打开 [cmd] 命令行 (macOS 请使用终端)，\n"
            f"复制并执行下方的自动修复指令将依赖装入插件目录中："
        )
        tk.Label(top_frame, text=info_text, justify=tk.LEFT, font=ui_font).pack(anchor=tk.W)
        
        # 生成基于当前 Sigil 水印插件所需包的专属 pip 安装命令
        cmd_str = f'pip install pyyaml Pillow resvg-py cffi imagequant --target="{str(_VENDOR_DIR)}"'
        
        text_box = tk.Text(top_frame, height=3, width=70, bg="#f5f6f7", font=code_font, relief=tk.FLAT)
        text_box.insert(tk.END, cmd_str)
        text_box.config(state=tk.DISABLED)
        text_box.pack(pady=15, fill=tk.X)
        
        def copy_cmd():
            guide_win.clipboard_clear()
            guide_win.clipboard_append(cmd_str)
            messagebox.showinfo("已复制", "命令已复制到剪贴板！\n\n请打开命令行界面右键粘贴并回车执行。\n安装完毕后重新启动本插件即可。", parent=guide_win)
            
        ttk.Button(top_frame, text="📋 一键复制修复指令", command=copy_cmd, padding=5).pack(pady=(5,0))
        
        guide_win.wait_window()
        err_root.destroy()
        return -1

    config_path = _PLUGIN_DIR / "watermarker_config.yaml"
    try:
        wm = SigilWatermarker(config_path)
    except Exception as e:
        print(f"【配置错误】{e}")
        return -1

    text_iter = list(bk.text_iter())
    if not text_iter:
        print("未找到任何 HTML 章节。")
        return 0

    root = tk.Tk()
    app = WatermarkApp(root, bk, wm, config_path)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    return 0

if __name__ == "__main__":
    print("请在 Sigil 中运行插件。")