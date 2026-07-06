#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import tkinter as tk
from tkinter import ttk, messagebox
from io import BytesIO
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser

# --- 1. 插件路径与第三方依赖加载 (方案三: 备胎策略) ---
_PLUGIN_DIR = Path(__file__).resolve().parent
_VENDOR_DIR = _PLUGIN_DIR / "vendor"

def setup_environment():
    """
    备胎策略：将 vendor 路径追加到 sys.path 末尾。
    优先让系统/Sigil 加载其自带的同名依赖，防止插件包破坏宿主环境。
    只有系统里完全没有这个包时，才会用到 vendor 里的内置包。
    """
    vendor_path = str(_VENDOR_DIR)
    if _VENDOR_DIR.is_dir():
        if vendor_path not in sys.path:
            sys.path.append(vendor_path)  # 核心改动：使用 append 而不是 insert(0)
            
    # Windows DLL 加载逻辑 (主要针对依赖 C 运行时的扩展包)
    if sys.platform == 'win32' and hasattr(os, 'add_dll_directory'):
        try:
            os.add_dll_directory(vendor_path)
            iq_dir = _VENDOR_DIR / "imagequant"
            if iq_dir.exists():
                os.add_dll_directory(str(iq_dir))
        except Exception:
            pass

setup_environment()

def _parse_version(v_str):
    """提取并解析版本号，例如 '9.0.1' -> (9, 0, 1)"""
    try:
        match = re.search(r'^(\d+\.\d+(\.\d+)?)', str(v_str))
        if match:
            return tuple(map(int, match.group(1).split('.')))
    except Exception:
        pass
    return (0, 0, 0)

def check_dependencies_and_versions():
    """
    严格检查依赖库。
    由于采用了 sys.path.append，极有可能加载的是宿主系统里的旧版本包，因此必须进行版本校验。
    """
    missing = []
    outdated = []

    # 检查 PyYAML (需 >= 5.1 支持 FullLoader)
    try: 
        import yaml
        if hasattr(yaml, '__version__') and _parse_version(yaml.__version__) < (5, 1):
            outdated.append(f"pyyaml (当前 {yaml.__version__}，需 >= 5.1)")
    except ImportError: 
        missing.append("pyyaml")

    # 检查 Pillow (建议 >= 9.0.0 以支持较新的 Resampling API)
    try: 
        import PIL
        from PIL import Image
        if hasattr(PIL, '__version__') and _parse_version(PIL.__version__) < (9, 0):
            outdated.append(f"Pillow (当前 {PIL.__version__}，需 >= 9.0.0)")
    except ImportError: 
        missing.append("Pillow")

    # 检查 resvg-py
    try: 
        from resvg_py import svg_to_bytes
    except ImportError: 
        missing.append("resvg-py")

    # 检查 cffi (明确告知是 imagequant 的依赖)
    try:
        import cffi
    except ImportError:
        missing.append("cffi (imagequant 的底层依赖)")

    # 检查 imagequant
    try: 
        import imagequant
    except ImportError: 
        missing.append("imagequant")

    return missing, outdated

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

# --- 2. 核心水印处理引擎 ---
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

    @property
    def watermark_buffer(self):
        return self._watermark_buffer.copy()

    def process_image_bytes(self, image_data: bytes) -> bytes:
        from PIL import Image
        import imagequant

        # 核心优化：全部操作基于内存流，极速合成
        with Image.open(BytesIO(image_data)).convert("RGBA") as img:
            iw, ih = img.size
            tw = iw * self.watermark_width // 100 if self.watermark_width_is_percent else self.watermark_width
            th = tw * self._watermark_buffer.height // self._watermark_buffer.width
            
            # 使用 getattr 兼容过低版本的 PIL (作为最后一道防线)
            resample_filter = getattr(Image.Resampling, 'LANCZOS', getattr(Image, 'LANCZOS', 1))
            wm_p = self.watermark_buffer.resize((tw, th), resample=resample_filter)
            
            if self.watermark_rotation != 0:
                wm_p = wm_p.rotate(self.watermark_rotation, expand=True)

            mx = iw * self.watermark_x_margin // 100 if self.watermark_x_margin_is_percent else self.watermark_x_margin
            my = ih * self.watermark_y_margin // 100 if self.watermark_y_margin_is_percent else self.watermark_y_margin
            px = mx if self.watermark_x_begin == "left" else iw - wm_p.width - mx
            py = my if self.watermark_y_begin == "top" else ih - wm_p.height - my

            # 内存/速度优化：仅针对水印区域修改 Alpha 通道
            if self.watermark_opacity < 1:
                alpha = wm_p.getchannel('A')
                alpha = alpha.point(lambda p: int(p * self.watermark_opacity))
                wm_p.putalpha(alpha)

            # 速度优化：直接利用水印的 alpha 作为掩码贴图，避免全尺寸图层运算
            img.paste(wm_p, (px, py), mask=wm_p)

            out_io = BytesIO()
            save_fmt = "JPEG" if self.output_format in ['jpg', 'jpeg'] else self.output_format.upper()
            if save_fmt == "JPEG": 
                img = img.convert("RGB")
            
            if self.output_quality < 100 and save_fmt == "PNG":
                imagequant.quantize_pil_image(img, max_quality=self.output_quality).save(out_io, format="PNG")
            else:
                img.save(out_io, format=save_fmt, quality=self.output_quality, optimize=True)
            
            return out_io.getvalue()

# --- 3. Python 原生 HTML 解析器引擎 ---
class NativeEpubHTMLParser(HTMLParser):
    def __init__(self, wm_config, img_map):
        super().__init__(convert_charrefs=True)
        self.exclude_classes = wm_config.exclude_classes
        self.exclude_tags = wm_config.exclude_tags
        self.exclude_images = wm_config.exclude_images
        self.image_map = img_map
        
        self.extracted_ids = set()
        self.skipped_ids = set()
        self.tag_stack = []
        
        self.current_hierarchy_classes = set()
        self.current_hierarchy_tags = set()

    def handle_starttag(self, tag, attrs):
        attr_dict = dict(attrs)
        tag = tag.lower()
        classes = set(attr_dict.get('class', '').split())
        
        self.tag_stack.append((tag, classes))
        self.current_hierarchy_tags.add(tag)
        self.current_hierarchy_classes.update(classes)
        
        self._check_image(tag, attr_dict)

    def handle_endtag(self, tag):
        tag = tag.lower()
        for i in range(len(self.tag_stack)-1, -1, -1):
            if self.tag_stack[i][0] == tag:
                del self.tag_stack[i:]
                self.current_hierarchy_tags = {t for t, _ in self.tag_stack}
                self.current_hierarchy_classes = set().union(*(cls for _, cls in self.tag_stack))
                break

    def handle_startendtag(self, tag, attrs):
        attr_dict = dict(attrs)
        tag = tag.lower()
        classes = set(attr_dict.get('class', '').split())
        
        temp_tags = self.current_hierarchy_tags | {tag}
        temp_classes = self.current_hierarchy_classes | classes
        self._check_image(tag, attr_dict, temp_tags, temp_classes)

    def _check_image(self, tag, attr_dict, current_tags=None, current_classes=None):
        if tag not in ('img', 'image'):
            return
            
        src = attr_dict.get('src') or attr_dict.get('xlink:href') or attr_dict.get('href')
        if not src: 
            return
        
        img_basename = src.split('/')[-1].split('#')[0].split('?')[0].lower()
        if img_basename not in self.image_map: 
            return
        
        img_id = self.image_map[img_basename]
        
        tags_to_check = current_tags if current_tags is not None else self.current_hierarchy_tags
        classes_to_check = current_classes if current_classes is not None else self.current_hierarchy_classes
        
        if (self.exclude_classes.intersection(classes_to_check) or 
            self.exclude_tags.intersection(tags_to_check) or 
            img_basename in self.exclude_images):
            self.skipped_ids.add(img_id)
        else:
            self.extracted_ids.add(img_id)

# --- 4. 彻底重构的 UI 交互与调度引擎 ---
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
        self.root.title("图片水印 - 章节选择")
        self.root.minsize(500, 480)
        self.root.eval('tk::PlaceWindow . center')

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1) 

        self.lbl_hint = ttk.Label(
            self.root, 
            text="注意：请在下方列表中，Ctrl选取或Shift选取。\n【二次确认】开启后，将允许你可视化核对并修改即将被处理的图片名单。", 
            foreground="red", 
            justify=tk.CENTER
        )
        self.lbl_hint.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        tree_height = max(8, min(len(self.text_iter), 15)) if self.text_iter else 8
        self.tree = ttk.Treeview(self.root, columns=("Data1", "Data2"), show="headings", height=tree_height, selectmode="extended")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Data1", anchor=tk.W, width=150, stretch=tk.YES)
        self.tree.column("Data2", anchor=tk.W, width=280, stretch=tk.YES)
        self.tree.heading("Data1", text="文件标识 (ID)")
        self.tree.heading("Data2", text="文件路径 (Href)")
        self.tree.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        scroll_bar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_bar.set)
        scroll_bar.grid(row=1, column=2, sticky=(tk.N, tk.S))

        for index, (data1, data2) in enumerate(self.text_iter):
            item_id = self.tree.insert(parent="", index=tk.END, iid=f"html_{index}", values=(data1, data2))
            if data1 in self.pre_selected_html_ids:
                self.tree.selection_add(item_id)

        self.btn_frame = ttk.Frame(self.root)
        self.btn_frame.grid(row=2, column=0, columnspan=3, pady=(10, 5))
        ttk.Button(self.btn_frame, text="全选", command=lambda: self.tree.selection_add(self.tree.get_children())).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="取消全选", command=lambda: self.tree.selection_remove(self.tree.get_children())).pack(side=tk.LEFT, padx=5)
        
        self.action_frame = ttk.Frame(self.root)
        self.action_frame.grid(row=3, column=0, columnspan=3, pady=(5, 10))
        
        self.btn_open_config = ttk.Button(self.action_frame, text="打开配置单", command=self.open_config_file)
        self.btn_open_config.pack(side=tk.LEFT, padx=5)

        self.secondary_confirm_var = tk.BooleanVar(value=False)
        self.chk_confirm = ttk.Checkbutton(self.action_frame, text="二次确认", variable=self.secondary_confirm_var)
        self.chk_confirm.pack(side=tk.LEFT, padx=5)
        
        self.btn_main = ttk.Button(self.action_frame, text="开始处理", command=self.run_action)
        self.btn_main.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(self.root, mode='determinate', length=400)
        self.progress.grid(row=4, column=0, columnspan=3, padx=10, pady=(10, 5), sticky=(tk.W, tk.E))
        self.progress_label = ttk.Label(self.root, text="等待开始...")
        self.progress_label.grid(row=5, column=0, columnspan=3, pady=(0, 10))

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
            btn_text = "正在处理中..."
        else:
            btn_text = "确认并执行" if self.phase == 2 else "开始处理"
            
        self.btn_main.config(state=state, text=btn_text)
        self.tree.config(selectmode='none' if state == tk.DISABLED else 'extended')
        
        chk_state = tk.NORMAL if (state == tk.NORMAL and self.phase == 1) else tk.DISABLED
        self.chk_confirm.config(state=chk_state)
        
        self.btn_open_config.config(state=state)
        for btn in self.btn_frame.winfo_children(): 
            btn.config(state=state)
        self.root.update()

    def process_images_batch(self, final_process_queue, skipped_image_ids):
        if not final_process_queue:
            messagebox.showinfo("完成", "未找到符合条件的图片，或图片均被拦截。")
            self.root.destroy()
            return

        self._set_ui_state(tk.DISABLED)
        self.progress['maximum'] = len(final_process_queue)
        self.progress['value'] = 0
        success_ids = []

        batch_size = max(4, self.wm.threads * 2) 

        with ThreadPoolExecutor(max_workers=self.wm.threads) as executor:
            for i in range(0, len(final_process_queue), batch_size):
                batch_uids = final_process_queue[i:i + batch_size]
                future_to_id = {}
                
                for uid in batch_uids:
                    image_data = self.bk.readfile(uid)
                    future = executor.submit(self.wm.process_image_bytes, image_data)
                    future_to_id[future] = uid
                
                for future in as_completed(future_to_id):
                    img_id = future_to_id[future]
                    try:
                        result_data = future.result()
                        self.bk.writefile(img_id, result_data)
                        success_ids.append(img_id)
                    except Exception as e:
                        print(f"渲染失败 {img_id}: {e}")
                        skipped_image_ids.add(img_id)
                    
                    self.progress['value'] += 1
                    self.progress_label.config(text=f"正在添加水印... {self.progress['value']}/{len(final_process_queue)}")
                    self.root.update()

        self._print_report(success_ids, skipped_image_ids)
        messagebox.showinfo("完成", f"批量水印添加完毕！\n成功: {len(success_ids)} 张\n跳过: {len(skipped_image_ids)} 张\n详细名单请查看 Sigil 控制台输出。")
        self.root.destroy()

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
            self.progress_label.config(text="正在解析 HTML 与黑名单...")
            
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
                self.root.title("图片水印 - 图片二次确认")
                self.lbl_hint.config(text="【二次确认】：下方为全书所有图片。即将被处理的图片已高亮选中，请按需调整后点击执行。")
                self.progress_label.config(text="请确认待处理图片...")
                
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

# --- 5. Sigil 插件入口 ---
def run(bk):
    missing, outdated = check_dependencies_and_versions()
    
    if missing or outdated:
        err_msg = "【依赖环境异常】\n为了防止干扰您的宿主环境，插件优先使用了您系统的本地包，但检测到异常：\n"
        
        if missing:
            err_msg += f"\n❌ 缺少库: {', '.join(missing)}"
        if outdated:
            err_msg += "\n\n⚠️ 版本过低（可能导致报错或崩溃）:"
            for pkg in outdated:
                err_msg += f"\n  - {pkg}"
                
        err_msg += ("\n\n【修复建议】\n请在您的命令行或终端中运行以下命令更新环境："
                    "\n\npip install --upgrade Pillow pyyaml resvg-py cffi imagequant")
        
        print("="*50)
        print(err_msg)
        print("="*50)
        
        tmp_root = tk.Tk()
        tmp_root.withdraw()
        messagebox.showerror("环境异常", err_msg)
        tmp_root.destroy()
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