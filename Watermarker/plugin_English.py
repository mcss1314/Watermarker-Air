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

# --- 1. Plugin path and third-party dependency loading (Strategy 3: Fallback) ---
_PLUGIN_DIR = Path(__file__).resolve().parent
_VENDOR_DIR = _PLUGIN_DIR / "vendor"

def setup_environment():
    """
    Fallback Strategy: Append the vendor path to the end of sys.path.
    Prioritize letting the system/Sigil load its own native dependencies 
    to prevent the plugin's packages from breaking the host environment.
    The built-in packages in the vendor directory will only be used if 
    the system completely lacks them.
    """
    vendor_path = str(_VENDOR_DIR)
    if _VENDOR_DIR.is_dir():
        if vendor_path not in sys.path:
            sys.path.append(vendor_path)  # Core change: Use append instead of insert(0)
            
    # Windows DLL loading logic (mainly for extension packages depending on C runtime)
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
    """Extract and parse the version number, e.g., '9.0.1' -> (9, 0, 1)"""
    try:
        match = re.search(r'^(\d+\.\d+(\.\d+)?)', str(v_str))
        if match:
            return tuple(map(int, match.group(1).split('.')))
    except Exception:
        pass
    return (0, 0, 0)

def check_dependencies_and_versions():
    """
    Strict dependency checking.
    Since sys.path.append is used, it's highly likely to load older packages 
    from the host system. Therefore, version validation is mandatory.
    """
    missing = []
    outdated = []

    # Check PyYAML (requires >= 5.1 to support FullLoader)
    try: 
        import yaml
        if hasattr(yaml, '__version__') and _parse_version(yaml.__version__) < (5, 1):
            outdated.append(f"pyyaml (current {yaml.__version__}, requires >= 5.1)")
    except ImportError: 
        missing.append("pyyaml")

    # Check Pillow (recommends >= 9.0.0 to support newer Resampling API)
    try: 
        import PIL
        from PIL import Image
        if hasattr(PIL, '__version__') and _parse_version(PIL.__version__) < (9, 0):
            outdated.append(f"Pillow (current {PIL.__version__}, requires >= 9.0.0)")
    except ImportError: 
        missing.append("Pillow")

    # Check resvg-py
    try: 
        from resvg_py import svg_to_bytes
    except ImportError: 
        missing.append("resvg-py")

    # Check cffi (explicitly state it is a dependency of imagequant)
    try:
        import cffi
    except ImportError:
        missing.append("cffi (underlying dependency of imagequant)")

    # Check imagequant
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
            raise UnidentifiedImageError(f"Failed to load watermark image: {e}")

# --- 2. Core Watermark Processing Engine ---
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
                raise PermissionError(f"Failed to write config file: {e}")
            raise ConfigError("Config file not found.\nA default config containing [Blacklist Rules] has been generated.\nPlease check the filter node and run again.")

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
            raise FileNotFoundError("Watermark image not found! Please ensure the path is correct.")

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

        # Core optimization: All operations based on memory stream, extremely fast composition
        with Image.open(BytesIO(image_data)).convert("RGBA") as img:
            iw, ih = img.size
            tw = iw * self.watermark_width // 100 if self.watermark_width_is_percent else self.watermark_width
            th = tw * self._watermark_buffer.height // self._watermark_buffer.width
            
            # Use getattr for compatibility with very old versions of PIL (as a last line of defense)
            resample_filter = getattr(Image.Resampling, 'LANCZOS', getattr(Image, 'LANCZOS', 1))
            wm_p = self.watermark_buffer.resize((tw, th), resample=resample_filter)
            
            if self.watermark_rotation != 0:
                wm_p = wm_p.rotate(self.watermark_rotation, expand=True)

            mx = iw * self.watermark_x_margin // 100 if self.watermark_x_margin_is_percent else self.watermark_x_margin
            my = ih * self.watermark_y_margin // 100 if self.watermark_y_margin_is_percent else self.watermark_y_margin
            px = mx if self.watermark_x_begin == "left" else iw - wm_p.width - mx
            py = my if self.watermark_y_begin == "top" else ih - wm_p.height - my

            # Memory/Speed optimization: Only modify the Alpha channel for the watermark area
            if self.watermark_opacity < 1:
                alpha = wm_p.getchannel('A')
                alpha = alpha.point(lambda p: int(p * self.watermark_opacity))
                wm_p.putalpha(alpha)

            # Speed optimization: Use the watermark's alpha directly as a mask to avoid full-size layer calculations
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

# --- 3. Python Native HTML Parser Engine ---
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

# --- 4. Completely refactored UI interaction and scheduling engine ---
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
        self.root.title("Image Watermark - Chapter Selection")
        self.root.minsize(500, 480)
        self.root.eval('tk::PlaceWindow . center')

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1) 

        self.lbl_hint = ttk.Label(
            self.root, 
            text="Note: Please select from the list below using Ctrl-click or Shift-click.\nWhen [Secondary Confirmation] is enabled, it allows you to visually verify and modify the list of images to be processed.", 
            foreground="red", 
            justify=tk.CENTER
        )
        self.lbl_hint.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        tree_height = max(8, min(len(self.text_iter), 15)) if self.text_iter else 8
        self.tree = ttk.Treeview(self.root, columns=("Data1", "Data2"), show="headings", height=tree_height, selectmode="extended")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Data1", anchor=tk.W, width=150, stretch=tk.YES)
        self.tree.column("Data2", anchor=tk.W, width=280, stretch=tk.YES)
        self.tree.heading("Data1", text="File ID")
        self.tree.heading("Data2", text="File Path (Href)")
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
        ttk.Button(self.btn_frame, text="Select All", command=lambda: self.tree.selection_add(self.tree.get_children())).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="Deselect All", command=lambda: self.tree.selection_remove(self.tree.get_children())).pack(side=tk.LEFT, padx=5)
        
        self.action_frame = ttk.Frame(self.root)
        self.action_frame.grid(row=3, column=0, columnspan=3, pady=(5, 10))
        
        self.btn_open_config = ttk.Button(self.action_frame, text="Open Config", command=self.open_config_file)
        self.btn_open_config.pack(side=tk.LEFT, padx=5)

        self.secondary_confirm_var = tk.BooleanVar(value=False)
        self.chk_confirm = ttk.Checkbutton(self.action_frame, text="Secondary Confirmation", variable=self.secondary_confirm_var)
        self.chk_confirm.pack(side=tk.LEFT, padx=5)
        
        self.btn_main = ttk.Button(self.action_frame, text="Start Processing", command=self.run_action)
        self.btn_main.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(self.root, mode='determinate', length=400)
        self.progress.grid(row=4, column=0, columnspan=3, padx=10, pady=(10, 5), sticky=(tk.W, tk.E))
        self.progress_label = ttk.Label(self.root, text="Waiting to start...")
        self.progress_label.grid(row=5, column=0, columnspan=3, pady=(0, 10))

    def open_config_file(self):
        import subprocess
        if not self.config_path.exists():
            messagebox.showwarning("Notice", "Config file does not exist, please ensure the environment is initialized.")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(self.config_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', str(self.config_path)])
            else:
                subprocess.call(['xdg-open', str(self.config_path)])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open config file: {e}")

    def _set_ui_state(self, state):
        if state == tk.DISABLED:
            btn_text = "Processing..."
        else:
            btn_text = "Confirm & Execute" if self.phase == 2 else "Start Processing"
            
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
            messagebox.showinfo("Complete", "No matching images found, or all images were intercepted.")
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
                        print(f"Render failed {img_id}: {e}")
                        skipped_image_ids.add(img_id)
                    
                    self.progress['value'] += 1
                    self.progress_label.config(text=f"Adding watermark... {self.progress['value']}/{len(final_process_queue)}")
                    self.root.update()

        self._print_report(success_ids, skipped_image_ids)
        messagebox.showinfo("Complete", f"Batch watermarking finished!\nSuccess: {len(success_ids)} images\nSkipped: {len(skipped_image_ids)} images\nPlease check the Sigil console output for the detailed list.")
        self.root.destroy()

    def _print_report(self, success_ids, skipped_image_ids):
        print("=" * 50)
        print("【Watermark Processing Report】")
        print("=" * 50)
        print(f"\n✅ Successfully processed files ({len(success_ids)}):")
        if success_ids:
            for sid in success_ids: print(f"  - {self.bk.id_to_href(sid).split('/')[-1]}")
        else: print("  (None)")

        print(f"\n🛑 Intercepted/Skipped files ({len(skipped_image_ids)}):")
        if skipped_image_ids:
            for xid in skipped_image_ids: print(f"  - {self.bk.id_to_href(xid).split('/')[-1]}")
        else: print("  (None)")
        print("=" * 50)

    def run_action(self):
        if self.phase == 1:
            selected_items = [self.tree.item(i)['values'] for i in self.tree.selection()]
            if not selected_items:
                messagebox.showwarning("Notice", "Please select the chapters to scan for images from the list first!")
                return

            self._set_ui_state(tk.DISABLED)
            self.progress_label.config(text="Parsing HTML and blacklists...")
            
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
                self.root.title("Image Watermark - Secondary Image Confirmation")
                self.lbl_hint.config(text="[Secondary Confirmation]: Below are all the images in the book. The images to be processed are highlighted. Please adjust as needed and click execute.")
                self.progress_label.config(text="Please confirm the images to be processed...")
                
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
                messagebox.showwarning("Notice", "No images selected for processing. Please reselect or close the window!")
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

# --- 5. Sigil Plugin Entry Point ---
def run(bk):
    missing, outdated = check_dependencies_and_versions()
    
    if missing or outdated:
        err_msg = "[Dependency Environment Exception]\nTo prevent interfering with your host environment, the plugin prioritized using your system's local packages, but detected exceptions:\n"
        
        if missing:
            err_msg += f"\n❌ Missing libraries: {', '.join(missing)}"
        if outdated:
            err_msg += "\n\n⚠️ Outdated versions (may cause errors or crashes):"
            for pkg in outdated:
                err_msg += f"\n  - {pkg}"
                
        err_msg += ("\n\n[Fix Suggestion]\nPlease run the following command in your terminal/command line to update the environment:\n\n"
                    "pip install --upgrade Pillow pyyaml resvg-py cffi imagequant")
        
        print("="*50)
        print(err_msg)
        print("="*50)
        
        tmp_root = tk.Tk()
        tmp_root.withdraw()
        messagebox.showerror("Environment Exception", err_msg)
        tmp_root.destroy()
        return -1

    config_path = _PLUGIN_DIR / "watermarker_config.yaml"
    try:
        wm = SigilWatermarker(config_path)
    except Exception as e:
        print(f"[Configuration Error] {e}")
        return -1

    text_iter = list(bk.text_iter())
    if not text_iter:
        print("No HTML chapters found.")
        return 0

    root = tk.Tk()
    app = WatermarkApp(root, bk, wm, config_path)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    return 0

if __name__ == "__main__":
    print("Please run this plugin within Sigil.")