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

# --- Plugin Path and Third-Party Dependencies Loading ---
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
            raise UnidentifiedImageError(f"Failed to load watermark image: {e}")

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
                raise PermissionError(f"Failed to write to configuration file: {e}")
            raise ConfigError("Configuration file does not exist.\nA default configuration containing [Blacklist Rules] has been generated.\nPlease check the filter node and run again.")

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

# --- Python Native HTML Parser Engine ---
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

# --- Sigil Plugin Entry Point ---
def run(bk):
    missing_deps = check_dependencies()
    if missing_deps:
        err_msg = f"[Error] Missing core libraries: {', '.join(missing_deps)}"
        if sys.platform != 'win32':
            err_msg += ("\n\n[Tip]\nThe built-in dependencies of this plugin only support Windows."
                        "\nNon-Windows users please manually install the missing dependencies:\npip install pyyaml Pillow resvg-py cffi imagequant")
        print(err_msg)
        tmp_root = tk.Tk()
        tmp_root.withdraw()
        messagebox.showerror("Missing Dependencies", err_msg)
        tmp_root.destroy()
        return -1

    config_path = _PLUGIN_DIR / "watermarker_config.yaml"
    try:
        wm = SigilWatermarker(config_path)
    except Exception as e:
        print(f"[Configuration Error] {e}")
        return -1

    all_image_map = {}
    for img_info in bk.image_iter():
        img_id, href = img_info[0], img_info[1]
        basename = href.split('/')[-1].split('#')[0].split('?')[0].lower()
        all_image_map[basename] = img_id

    text_iter = list(bk.text_iter())
    if not text_iter:
        print("No HTML chapters found.")
        return 0

    pre_selected_html_ids = set()
    try:
        for id_type, Id in bk.selected_iter():
            if id_type == "text": pre_selected_html_ids.add(Id)
    except Exception:
        pass

    # ================= UI and State Control Logic =================
    app_state = {"phase": 1}  # Phase 1: Select HTML, Phase 2: Secondary confirmation of images

    root = tk.Tk()
    root.title("Image Watermark - Chapter Selection")
    root.minsize(500, 480)
    root.eval('tk::PlaceWindow . center')

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(1, weight=1) 

    lbl_hint = ttk.Label(
        root, 
        text="Note: Please use Ctrl or Shift to select in the list below.\nAfter enabling [Secondary Confirmation], you can visually verify and modify the list of images to be processed.", 
        foreground="red", 
        justify=tk.CENTER
    )
    lbl_hint.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

    tree_height = max(8, min(len(text_iter), 15))
    tree = ttk.Treeview(root, columns=("Data1", "Data2"), show="headings", height=tree_height, selectmode="extended")
    tree.column("#0", width=0, stretch=tk.NO)
    tree.column("Data1", anchor=tk.W, width=150, stretch=tk.YES)
    tree.column("Data2", anchor=tk.W, width=280, stretch=tk.YES)
    tree.heading("Data1", text="File Identifier (ID)")
    tree.heading("Data2", text="File Path (Href)")
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
    ttk.Button(btn_frame, text="Select All", command=lambda: tree.selection_add(tree.get_children())).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Deselect All", command=lambda: tree.selection_remove(tree.get_children())).pack(side=tk.LEFT, padx=5)
    
    action_frame = ttk.Frame(root)
    action_frame.grid(row=3, column=0, columnspan=3, pady=(5, 10))
    
    def open_config_file():
        if not config_path.exists():
            messagebox.showwarning("Notice", "Configuration file does not exist. Please ensure the environment is initialized.")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(config_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', str(config_path)])
            else:
                subprocess.call(['xdg-open', str(config_path)])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open configuration file: {e}")

    btn_open_config = ttk.Button(action_frame, text="Open Configuration", command=open_config_file)
    btn_open_config.pack(side=tk.LEFT, padx=5)

    secondary_confirm_var = tk.BooleanVar(value=False)
    chk_confirm = ttk.Checkbutton(action_frame, text="Secondary Confirmation", variable=secondary_confirm_var)
    chk_confirm.pack(side=tk.LEFT, padx=5)
    
    button = ttk.Button(action_frame, text="Start Processing")
    button.pack(side=tk.LEFT, padx=5)

    progress = ttk.Progressbar(root, mode='determinate', length=400)
    progress.grid(row=4, column=0, columnspan=3, padx=10, pady=(10, 5), sticky=(tk.W, tk.E))
    progress_label = ttk.Label(root, text="Waiting to start...")
    progress_label.grid(row=5, column=0, columnspan=3, pady=(0, 10))

    def _set_ui_state(state):
        if state == tk.DISABLED:
            btn_text = "Processing..."
        else:
            btn_text = "Confirm and Execute" if app_state["phase"] == 2 else "Start Processing"
            
        button.config(state=state, text=btn_text)
        tree.config(selectmode='none' if state == tk.DISABLED else 'extended')
        
        # Secondary confirmation checkbox logic: only available in Phase 1 and when fully enabled, otherwise grayed out
        chk_state = tk.NORMAL if (state == tk.NORMAL and app_state["phase"] == 1) else tk.DISABLED
        chk_confirm.config(state=chk_state)
        
        btn_open_config.config(state=state)
        for btn in btn_frame.winfo_children(): 
            btn.config(state=state)

    def execute_watermark(final_process_queue, skipped_image_ids):
        if not final_process_queue:
            messagebox.showinfo("Complete", "No eligible images found, or all images were intercepted.")
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
                    print(f"Rendering failed {img_id}: {e}")
                    skipped_image_ids.add(img_id)
                
                progress['value'] += 1
                progress_label.config(text=f"Adding watermark... {progress['value']}/{len(final_process_queue)}")
                root.update()

        # Print final report
        print("=" * 50)
        print("[Watermark Processing Report]")
        print("=" * 50)
        print(f"\n✅ Successfully processed files ({len(success_ids)}):")
        if success_ids:
            for sid in success_ids: print(f"  - {bk.id_to_href(sid).split('/')[-1]}")
        else: print("  (None)")

        print(f"\n🛑 Intercepted/Skipped files ({len(skipped_image_ids)}):")
        if skipped_image_ids:
            for xid in skipped_image_ids: print(f"  - {bk.id_to_href(xid).split('/')[-1]}")
        else: print("  (None)")
        print("=" * 50)

        messagebox.showinfo("Complete", f"Batch watermarking completed!\nSuccess: {len(success_ids)} images\nSkipped: {len(skipped_image_ids)} images\nPlease check the Sigil console output for the detailed list.")
        root.destroy()

    def run_action():
        if app_state["phase"] == 1:
            selected_items = [tree.item(i)['values'] for i in tree.selection()]
            if not selected_items:
                messagebox.showwarning("Notice", "Please first select the chapters in the list that need to be scanned for images!")
                return

            _set_ui_state(tk.DISABLED)
            progress_label.config(text="Parsing HTML and blacklist...")
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
                # ===== Enter Phase 2: Image Secondary Confirmation Mode =====
                app_state["phase"] = 2
                root.title("Image Watermark - Image Secondary Confirmation")
                lbl_hint.config(text="[Secondary Confirmation]: Below are all the images in the book. The images to be processed are highlighted. Please adjust as needed and click Execute.")
                progress_label.config(text="Please confirm the images to be processed...")
                
                # Clear Treeview and reload all book images
                for item in tree.get_children():
                    tree.delete(item)
                
                for index, img_info in enumerate(bk.image_iter()):
                    img_id, href = img_info[0], img_info[1]
                    href_display = href if href else ""
                    item_id = tree.insert(parent="", index=tk.END, iid=f"img_{index}", values=(img_id, href_display))
                    # Only images in the final queue will be selected by default
                    if img_id in final_process_queue:
                        tree.selection_add(item_id)
                
                _set_ui_state(tk.NORMAL) 
                return
            else:
                # Directly enter processing
                execute_watermark(final_process_queue, skipped_image_ids)

        elif app_state["phase"] == 2:
            # ===== Phase 2: Execute selected images =====
            selected_items = [tree.item(i)['values'] for i in tree.selection()]
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
                    if ext in wm.target_formats:
                        final_process_queue.append(img_id)
                    else:
                        skipped_image_ids.add(img_id)
                else:
                    # Fallback attempt, let PIL handle error throwing
                    final_process_queue.append(img_id)
                    
            execute_watermark(final_process_queue, skipped_image_ids)

    button.config(command=run_action)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    return 0

if __name__ == "__main__":
    print("Please run the plugin within Sigil.")