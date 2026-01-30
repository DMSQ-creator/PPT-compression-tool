# 🚀 PPT 极致压缩工具 (PPT Ultra Compressor)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue) ![GUI](https://img.shields.io/badge/GUI-ttkbootstrap-success) ![License](https://img.shields.io/badge/License-MIT-green)

一款基于 Python 开发的本地化、轻量级、高颜值的 PowerPoint (`.pptx`) 文件压缩工具。
**无需联网**，保护隐私，通过智能算法大幅减小 PPT 体积，解决微信/邮件无法发送超大 PPT 的烦恼。

![软件界面截图](https://github.com/DMSQ-creator/PPT-compression-tool/blob/main/image.png)

## ✨ 核心亮点

*   **🔒 本地运行，隐私安全**：所有操作均在本地完成，文件无需上传到第三方服务器，绝无泄露风险。
*   **🖼️ 智能图像算法**：
    *   **PNG 智能量化**：引入色彩量化（Color Quantization）技术，将 32位 真彩 PNG 转换为 8位 索引图，在保留透明度的前提下减少 60%-80% 体积。
    *   **Lanczos 抗锯齿**：使用高质量重采样滤镜，确保图片缩小后依然清晰锐利。
*   **⚡ 极速体验**：采用多线程架构，压缩大文件时界面丝滑流畅，配备高清抗锯齿加载动画。
*   **🎨 现代化 UI**：基于 `ttkbootstrap` 构建的扁平化界面，简洁美观（Cosmo 主题）。
*   **🛡️ 无损备份**：自动生成 `文件名_高清压缩.pptx`，绝不覆盖原文件，安全无忧。

## ⚙️ 压缩模式说明

| 模式 | 适用场景 | 技术参数 |
| :--- | :--- | :--- |
| **🏆 智能高清 (推荐)** | 只有超大图被压缩，肉眼几乎无损 | Max Width: 2048px / Quality: 85 |
| **⚖️ 均衡模式** | 适合日常办公、传阅 | Max Width: 1600px / Quality: 75 |
| **📉 强力压缩** | 适合手机查看、微信发送 | Max Width: 1280px / Quality: 60 + **PNG量化** |
| **🔥 极限压缩** | 只要能看清字就行，极致体积 | Max Width: 1024px / Quality: 50 + **PNG量化** |

## 📥 下载与安装

### 方式一：直接运行 (Windows)
请前往 [Releases页面](https://github.com/DMSQ-creator/PPT-compression-tool/releases) 下载最新的绿色版文件夹。
1. 下载压缩包并解压。
2. 进入文件夹，双击 `PPT极致压缩工具.exe` 即可秒开使用。

### 方式二：源码运行
如果您是开发者，可以克隆仓库运行：

```bash
# 1. 克隆仓库
git clone https://github.com/your-username/ppt-compressor.git
cd ppt-compressor

# 2. 安装依赖
pip install ttkbootstrap pillow

# 3. 运行
python ppt_compressor.py
```

## 🛠️ 开发与打包指南

本项目使用 `PyInstaller` 进行打包。为了获得最快的启动速度，推荐使用**文件夹模式 (Onedir)**。

### 1. 准备环境
建议在纯净的虚拟环境中打包，以减小体积：
```bash
python -m venv venv
venv\Scripts\activate
pip install ttkbootstrap pillow pyinstaller
```

### 2. 执行打包命令
```bash
pyinstaller --noconfirm --onedir --windowed --clean --icon="app.ico" --name="PPT极致压缩工具" --collect-all ttkbootstrap ppt_compressor.py
```

*   `--onedir`: 生成文件夹，启动速度极快。
*   `--windowed`: 隐藏控制台黑窗口。
*   `--collect-all ttkbootstrap`: 确保包含所有 UI 主题资源。

## 🧠 技术原理

1.  **解压结构**：`.pptx` 文件本质上是一个 Zip 压缩包。程序将其解压并读取内部结构。
2.  **定位媒体**：遍历 `ppt/media/` 目录，找到所有图片资源。
3.  **智能处理**：
    *   对于 **JPEG**：重新编码并调整质量 (Quality)。
    *   对于 **PNG**：如果选择强力模式，使用 `img.quantize(colors=256)` 进行色彩缩减。
    *   对于 **尺寸**：如果图片分辨率超过阈值，使用 Lanczos 算法进行缩放。
4.  **重组打包**：将处理后的图片流直接写入新的 Zip 结构中，修改后缀为 `.pptx`。

## ⚠️ 注意事项

*   目前的版本主要针对**图片**进行压缩。如果 PPT 中包含大量高清**视频**或**音频**文件，建议使用专门的视频压缩工具处理后再插入。
*   矢量图（SVG/EMF/WMF）不会被处理。

## 📜 开源协议

MIT License. 欢迎 Fork 和 Star！🌟
