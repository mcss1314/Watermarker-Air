# Watermarker-Air

This plugin is developed based on the Sigil plugin mechanism. It can batch-add watermarks to images in EPUB ebooks and supports excluding images that do not require watermarks via blacklist rules. The plugin provides a graphical configuration interface, allowing users to customize blacklist rules and perform secondary confirmation on files to be processed.

## 🌟 Core Features

- **Multi-format Compatibility**: Supports reading and processing JPG, JPEG, PNG, and WebP image formats.
- **Vector and Bitmap Watermark Sources**: Supports not only standard raster images as watermarks but also parses SVG files (rendered into high-quality graphic buffers via the underlying `resvg-py` library).
- **Contextual Smart Filtering**: Features a robust blacklist defense mechanism. By inheriting from `HTMLParser` to track the tag stack, the plugin can automatically skip specific HTML files (like `cover.xhtml`), specific image names, containers with specific CSS class names (like `no-watermark`), and images within specific HTML tags.
- **Two-Phase Graphical User Interface (GUI)**:
  - **Chapter Selector**: Phase 1 automatically lists all internal HTML files and intelligently pre-selects the chapters you are currently editing in the Sigil main interface.
  - **Image Secondary Confirmation**: Provides an optional secondary confirmation. When enabled, the system scans the selected chapters and visually displays all matched images in Phase 2, allowing for precise single-image exclusion.
- **Multi-threading Acceleration**: Utilizes `ThreadPoolExecutor` under the hood to build a concurrent task pool, with 4 rendering threads enabled by default, significantly reducing processing time for EPUB projects with a massive number of images.
- **Size Optimization and Quantization**: Built-in smart compression strategy. When a specific output quality is specified, it automatically calls the `imagequant` library to perform color quantization compression on PNG images, effectively preventing ebook bloat caused by adding watermarks.

## 📦 Environment Dependencies

Before running this plugin, ensure that the following third-party dependencies are installed in your external Python environment for Sigil or within the plugin's bundled `vendor` directory:

- `Pillow` (Core image processing engine)
- `pyyaml` (Configuration file parsing)
- `resvg-py` (SVG vector file rasterization rendering)
- `imagequant` (Advanced lossy compression optimization for PNG)
- `cffi` (Provides a C foreign function interface for Python)

## 🚀 Installation Guide

1. In Sigil, go to the menu bar `Plugins` → `Manage Plugins` to add the plugin.
2. After successful installation, the plugin is located in Sigil's official plugin directory. For Windows systems, the default path is usually `…\sigil-ebook\sigil\plugins`.
3. Restart Sigil, and you will see and be able to execute this tool from the `Plugins` menu at the top.

## ⚙️ Configuration Instructions

Upon first initialization, if no configuration is detected, the plugin automatically generates a default `watermarker_config.yaml` file in the current directory. You can freely customize the following parameters according to your project needs:

- **Watermark Attributes**: Configure `image_filename` to set the watermark source file; use `width` to set the watermark size (supports fixed pixels or a percentage relative to the original image); adjust `opacity` to set the transparency blend level; and configure `rotation` to change the watermark's tilt angle.
- **Positioning Margins**: Through the `margins` parameters for the X and Y axes (also supporting pixels or percentages), you can precisely anchor the watermark to any corner of the image.
- **Quality Control**: Control the output rendering quality with a value between `0` and `100`.
- **Filter Rule List**: Declare specific CSS class names or file identifiers in the configuration file that you wish to globally skip during processing.

## 💡 Usage Workflow

1. Open the EPUB ebook you want to process in Sigil.
2. Find the `Plugins` menu and launch the **Watermarker-Air** plugin.
3. In the pop-up Phase 1 GUI window, check the chapters where you need to batch inject watermarks.
4. If you checked "Secondary Confirmation", review and fine-tune the actual list of images to be processed in the subsequent Phase 2 window.
5. After clicking execute, the plugin will start multi-threaded rendering in the background. Once completed, you can open Sigil's **Plugin Console** to view a detailed execution report, which will list all successfully composited image records as well as skipped image records triggered by filter rules.