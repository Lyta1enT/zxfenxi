# 征信报告OCR识别与生成工具

支持上传 PDF、图片、Word、Excel，自动识别并汇总征信信息，生成结构化 Excel 报告（10列客户格式）。

## 快速使用

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动
python main.py
```

## 打包成 Windows 软件

### 在 Windows 上打包

1. **安装 Python 3.8+**（https://www.python.org/downloads/）
2. 双击 `build.bat`，等待打包完成
3. 输出文件：`dist/征信报告OCR工具.exe`

### 手动打包

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包（单文件，隐藏控制台窗口）
pyinstaller --clean --windowed --onefile --name "征信报告OCR工具" main.py

# 输出在 dist/ 目录
```

### 跨平台打包（在 macOS 上打包 Windows 版本）

```bash
pip install pyinstaller
# 需要安装 wine 或使用 Docker 的 Windows 容器
```

## 注意事项

- 首次运行需要下载 EasyOCR 模型文件（约 3-5 分钟）
- 模型缓存后不再重复下载
- OCR 识别在旧型号 CPU 上可能较慢（5-30秒/张图片）
