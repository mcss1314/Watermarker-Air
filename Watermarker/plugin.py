#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from io import BytesIO
from pathlib import Path
from traceback import format_exc
from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser

# --- 插件路径与第三方依赖加载 ---
_PLUGIN_DIR = Path(__file__).resolve().parent
_VENDOR_DIR = _PLUGIN_DIR / "vendor"

def setup_environment():
    if _VENDOR_DIR.is_dir() and str(_VENDOR_DIR) not in sys.path:
        sys.path.insert(0, str(_VENDOR_DIR))
    if sys.platform == 'win32' and hasattr(os, 'add_dll_directory'):
        try:
            os.add_dll_directory(str(_VENDOR_DIR))
            iq_dir = _VENDOR_DIR / "imagequant"
            if iq_dir.exists():
                os.add_dll_directory(str(iq_dir))
        except Exception:
            pass

setup_environment()

def check_dependencies():
    missing = []
    try: import yaml
    except ImportError: missing.append("pyyaml")
    try: from PIL import Image
    except ImportError: missing.append("Pillow")
    try: from resvg_py import svg_to_bytes
    except ImportError: missing.append("resvg-py")
    try: import imagequant
    except ImportError: missing.append("imagequant")
    return missing

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
            "image_filename": "plugin.png",
            "width": {"unit_is_px_or_percent": "percent", "width_value": 10},
            "x_margin": {"begin_from_left_or_right": "right", "unit_is_px_or_percent": "percent", "margin_value": 5},
            "y_margin": {"begin_from_top_or_bottom": "bottom", "unit_is_px_or_percent": "percent", "margin_value": 5},
            "opacity": 0.5,
            "rotation": 0,
        },
    }

    def __init__(self, config_yaml_path: Path, max_epub_img_width: int = 1000) -> None:
        from yaml import load, dump, FullLoader
        from PIL import Image

        if not config_yaml_path.exists():
            try:
                with open(config_yaml_path, "w", encoding="utf-8") as f:
                    dump(self.default, f, allow_unicode=True, sort_keys=False)
            except Exception as e:
                raise PermissionError(f"无法写入配置文件：{e}")
            raise ConfigError("配置文件不存在。\n已生成包含【黑名单规则】的默认配置。\n请检查 filter 节点后重新运行。")

        with open(config_yaml_path, "r", encoding="utf-8") as f:
            data = load(f, Loader=FullLoader)

        self.threads = data.get("threads", 4)
        filter_data = data.get("filter", {})
        
        self.exclude_html_stems = {Path(n).stem.lower() for n in (filter_data.get("exclude_html") or [])}
        self.exclude_images = [n.lower() for n in (filter_data.get("exclude_images") or [])]
        self.exclude_classes = set(filter_data.get("exclude_classes") or [])
        self.exclude_tags = set(n.lower() for n in (filter_data.get("exclude_tags") or []))

        process_data = data.get("process", {})
        self.target_formats = [fmt.lower() for fmt in process_data.get("target_formats", ["jpg", "jpeg", "png"])]
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

        with Image.open(BytesIO(image_data)).convert("RGBA") as img:
            iw, ih = img.size
            tw = iw * self.watermark_width // 100 if self.watermark_width_is_percent else self.watermark_width
            th = tw * self._watermark_buffer.height // self._watermark_buffer.width
            
            wm_p = self.watermark_buffer.resize((tw, th), resample=Image.Resampling.LANCZOS)
            if self.watermark_rotation != 0:
                wm_p = wm_p.rotate(self.watermark_rotation, expand=True)

            mx = iw * self.watermark_x_margin // 100 if self.watermark_x_margin_is_percent else self.watermark_x_margin
            my = ih * self.watermark_y_margin // 100 if self.watermark_y_margin_is_percent else self.watermark_y_margin
            px = mx if self.watermark_x_begin == "left" else iw - wm_p.width - mx
            py = my if self.watermark_y_begin == "top" else ih - wm_p.height - my

            wm_layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
            wm_layer.paste(wm_p, (px, py), wm_p)
            if self.watermark_opacity < 1:
                alpha = wm_layer.split()[3].point(lambda p: int(p * self.watermark_opacity))
                wm_layer.putalpha(alpha)

            final = Image.alpha_composite(img, wm_layer)

            out_io = BytesIO()
            save_fmt = "JPEG" if self.output_format in ['jpg', 'jpeg'] else self.output_format.upper()
            if save_fmt == "JPEG": final = final.convert("RGB")
            
            if self.output_quality < 100 and save_fmt == "PNG":
                imagequant.quantize_pil_image(final, max_quality=self.output_quality).save(out_io, format="PNG")
            else:
                final.save(out_io, format=save_fmt, quality=self.output_quality, optimize=True)
            
            return out_io.getvalue()

# --- Python 原生 HTML 解析器引擎 ---
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

    def handle_starttag(self, tag, attrs):
        attr_dict = dict(attrs)
        tag = tag.lower()
        classes = set(attr_dict.get('class', '').split())
        self.tag_stack.append((tag, classes))
        self._check_image(tag, attr_dict)

    def handle_endtag(self, tag):
        tag = tag.lower()
        for i in range(len(self.tag_stack)-1, -1, -1):
            if self.tag_stack[i][0] == tag:
                del self.tag_stack[i:]
                break

    def handle_startendtag(self, tag, attrs):
        attr_dict = dict(attrs)
        tag = tag.lower()
        classes = set(attr_dict.get('class', '').split())
        self.tag_stack.append((tag, classes))
        self._check_image(tag, attr_dict)
        self.tag_stack.pop()

    def _check_image(self, tag, attr_dict):
        if tag in ('img', 'image'):
            src = attr_dict.get('src') or attr_dict.get('xlink:href') or attr_dict.get('href')
            if not src: return
            
            img_basename = src.split('/')[-1].split('#')[0].split('?')[0].lower()
            if img_basename not in self.image_map: return
            
            img_id = self.image_map[img_basename]

            hierarchy_classes = set()
            hierarchy_tags = set()
            for t, parent_classes in self.tag_stack:
                hierarchy_tags.add(t)
                hierarchy_classes.update(parent_classes)
            
            if self.exclude_classes.intersection(hierarchy_classes) or \
               self.exclude_tags.intersection(hierarchy_tags) or \
               img_basename in self.exclude_images:
                self.skipped_ids.add(img_id)
                return
                
            self.extracted_ids.add(img_id)

# --- Sigil 插件入口 ---
def run(bk):
    missing_deps = check_dependencies()
    if missing_deps:
        err_msg = f"【错误】缺少核心库：{', '.join(missing_deps)}"
        if sys.platform != 'win32':
            err_msg += ("\n\n【温馨提示】\n当前插件内置的依赖包仅支持 Windows 系统。"
                        "\n非 Win 用户请手动安装缺失依赖：\npip install pyyaml Pillow resvg-py cffi imagequant")
        print(err_msg)
        tmp_root = tk.Tk()
        tmp_root.withdraw()
        messagebox.showerror("依赖缺失", err_msg)
        tmp_root.destroy()
        return -1

    config_path = _PLUGIN_DIR / "watermarker_config.yaml"
    try:
        wm = SigilWatermarker(config_path)
    except Exception as e:
        print(f"【配置错误】{e}")
        return -1

    all_image_map = {}
    for img_info in bk.image_iter():
        img_id, href = img_info[0], img_info[1]
        basename = href.split('/')[-1].split('#')[0].split('?')[0].lower()
        all_image_map[basename] = img_id

    text_iter = list(bk.text_iter())
    if not text_iter:
        print("未找到任何 HTML 章节。")
        return 0

    pre_selected_html_ids = set()
    try:
        for id_type, Id in bk.selected_iter():
            if id_type == "text": pre_selected_html_ids.add(Id)
    except Exception:
        pass

    # ================= UI 界面与状态控制逻辑 =================
    app_state = {"phase": 1}  # Phase 1: 选 HTML, Phase 2: 二次确认图片

    root = tk.Tk()
    root.title("图片水印 - 章节选择")
    root.minsize(500, 480)
    root.eval('tk::PlaceWindow . center')

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(1, weight=1) 

    lbl_hint = ttk.Label(
        root, 
        text="注意：请在下方列表中，Ctrl选取或Shift选取。\n【二次确认】开启后，将允许你可视化核对并修改即将被处理的图片名单。", 
        foreground="red", 
        justify=tk.CENTER
    )
    lbl_hint.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

    tree_height = max(8, min(len(text_iter), 15))
    tree = ttk.Treeview(root, columns=("Data1", "Data2"), show="headings", height=tree_height, selectmode="extended")
    tree.column("#0", width=0, stretch=tk.NO)
    tree.column("Data1", anchor=tk.W, width=150, stretch=tk.YES)
    tree.column("Data2", anchor=tk.W, width=280, stretch=tk.YES)
    tree.heading("Data1", text="文件标识 (ID)")
    tree.heading("Data2", text="文件路径 (Href)")
    tree.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

    scroll_bar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scroll_bar.set)
    scroll_bar.grid(row=1, column=2, sticky=(tk.N, tk.S))

    for index, (data1, data2) in enumerate(text_iter):
        item_id = tree.insert(parent="", index=tk.END, iid=f"html_{index}", values=(data1, data2))
        if data1 in pre_selected_html_ids:
            tree.selection_add(item_id)

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=2, column=0, columnspan=3, pady=(10, 5))
    ttk.Button(btn_frame, text="全选", command=lambda: tree.selection_add(tree.get_children())).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="取消全选", command=lambda: tree.selection_remove(tree.get_children())).pack(side=tk.LEFT, padx=5)
    
    action_frame = ttk.Frame(root)
    action_frame.grid(row=3, column=0, columnspan=3, pady=(5, 10))
    
    def open_config_file():
        if not config_path.exists():
            messagebox.showwarning("提示", "配置文件不存在，请确保已初始化环境。")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(config_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', str(config_path)])
            else:
                subprocess.call(['xdg-open', str(config_path)])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开配置文件：{e}")

    btn_open_config = ttk.Button(action_frame, text="打开配置单", command=open_config_file)
    btn_open_config.pack(side=tk.LEFT, padx=5)

    secondary_confirm_var = tk.BooleanVar(value=False)
    chk_confirm = ttk.Checkbutton(action_frame, text="二次确认", variable=secondary_confirm_var)
    chk_confirm.pack(side=tk.LEFT, padx=5)
    
    button = ttk.Button(action_frame, text="开始处理")
    button.pack(side=tk.LEFT, padx=5)

    progress = ttk.Progressbar(root, mode='determinate', length=400)
    progress.grid(row=4, column=0, columnspan=3, padx=10, pady=(10, 5), sticky=(tk.W, tk.E))
    progress_label = ttk.Label(root, text="等待开始...")
    progress_label.grid(row=5, column=0, columnspan=3, pady=(0, 10))

    def _set_ui_state(state):
        if state == tk.DISABLED:
            btn_text = "正在处理中..."
        else:
            btn_text = "确认并执行" if app_state["phase"] == 2 else "开始处理"
            
        button.config(state=state, text=btn_text)
        tree.config(selectmode='none' if state == tk.DISABLED else 'extended')
        
        # 二次确认勾选框逻辑：仅在 Phase 1 且整体开启时可用，否则全灰
        chk_state = tk.NORMAL if (state == tk.NORMAL and app_state["phase"] == 1) else tk.DISABLED
        chk_confirm.config(state=chk_state)
        
        btn_open_config.config(state=state)
        for btn in btn_frame.winfo_children(): 
            btn.config(state=state)

    def execute_watermark(final_process_queue, skipped_image_ids):
        if not final_process_queue:
            messagebox.showinfo("完成", "未找到符合条件的图片，或图片均被拦截。")
            root.destroy()
            return

        _set_ui_state(tk.DISABLED)
        progress['maximum'] = len(final_process_queue)
        progress['value'] = 0
        success_ids = []

        with ThreadPoolExecutor(max_workers=wm.threads) as executor:
            future_to_id = {executor.submit(wm.process_image_bytes, bk.readfile(uid)): uid for uid in final_process_queue}
            
            for future in as_completed(future_to_id):
                img_id = future_to_id[future]
                try:
                    result_data = future.result()
                    bk.writefile(img_id, result_data)
                    success_ids.append(img_id)
                except Exception as e:
                    print(f"渲染失败 {img_id}: {e}")
                    skipped_image_ids.add(img_id)
                
                progress['value'] += 1
                progress_label.config(text=f"正在添加水印... {progress['value']}/{len(final_process_queue)}")
                root.update()

        # 打印最终报表
        print("=" * 50)
        print("【水印处理报告】")
        print("=" * 50)
        print(f"\n✅ 处理成功的文件 ({len(success_ids)}个):")
        if success_ids:
            for sid in success_ids: print(f"  - {bk.id_to_href(sid).split('/')[-1]}")
        else: print("  (无)")

        print(f"\n🛑 拦截/略过的文件 ({len(skipped_image_ids)}个):")
        if skipped_image_ids:
            for xid in skipped_image_ids: print(f"  - {bk.id_to_href(xid).split('/')[-1]}")
        else: print("  (无)")
        print("=" * 50)

        messagebox.showinfo("完成", f"批量水印添加完毕！\n成功: {len(success_ids)} 张\n跳过: {len(skipped_image_ids)} 张\n详细名单请查看 Sigil 控制台输出。")
        root.destroy()

    def run_action():
        if app_state["phase"] == 1:
            selected_items = [tree.item(i)['values'] for i in tree.selection()]
            if not selected_items:
                messagebox.showwarning("提示", "请先在列表中选择需要扫描配图的章节！")
                return

            _set_ui_state(tk.DISABLED)
            progress_label.config(text="正在解析 HTML 与黑名单...")
            root.update()

            selected_html_ids = [item[0] for item in selected_items]
            target_image_ids = set()
            skipped_image_ids = set()

            for html_id in selected_html_ids:
                html_data = bk.readfile(html_id)
                parser = NativeEpubHTMLParser(wm, all_image_map)
                parser.feed(html_data)
                
                href = bk.id_to_href(html_id)
                html_stem = Path(href.split('/')[-1].split('#')[0].split('?')[0]).stem.lower()
                
                if html_stem in wm.exclude_html_stems:
                    skipped_image_ids.update(parser.extracted_ids)
                    skipped_image_ids.update(parser.skipped_ids)
                else:
                    target_image_ids.update(parser.extracted_ids)
                    skipped_image_ids.update(parser.skipped_ids)

            target_image_ids.difference_update(skipped_image_ids)

            final_process_queue = []
            for img_id in target_image_ids.copy():
                ext = bk.id_to_href(img_id).rsplit('.')[-1].lower()
                if ext in wm.target_formats:
                    final_process_queue.append(img_id)
                else:
                    skipped_image_ids.add(img_id)

            if secondary_confirm_var.get():
                # ===== 进入 Phase 2: 图片二次确认模式 =====
                app_state["phase"] = 2
                root.title("图片水印 - 图片二次确认")
                lbl_hint.config(text="【二次确认】：下方为全书所有图片。即将被处理的图片已高亮选中，请按需调整后点击执行。")
                progress_label.config(text="请确认待处理图片...")
                
                # 清空 Treeview 并重载全书图片
                for item in tree.get_children():
                    tree.delete(item)
                
                for index, img_info in enumerate(bk.image_iter()):
                    img_id, href = img_info[0], img_info[1]
                    href_display = href if href else ""
                    item_id = tree.insert(parent="", index=tk.END, iid=f"img_{index}", values=(img_id, href_display))
                    # 只有在最终队列中的图片才会被默认选中
                    if img_id in final_process_queue:
                        tree.selection_add(item_id)
                
                _set_ui_state(tk.NORMAL) 
                return
            else:
                # 直接进入处理
                execute_watermark(final_process_queue, skipped_image_ids)

        elif app_state["phase"] == 2:
            # ===== Phase 2: 执行选中的图片 =====
            selected_items = [tree.item(i)['values'] for i in tree.selection()]
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
                    if ext in wm.target_formats:
                        final_process_queue.append(img_id)
                    else:
                        skipped_image_ids.add(img_id)
                else:
                    # 保底尝试，交由 PIL 处理抛错判断
                    final_process_queue.append(img_id)
                    
            execute_watermark(final_process_queue, skipped_image_ids)

    button.config(command=run_action)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    return 0

if __name__ == "__main__":
    print("请在 Sigil 中运行插件。")