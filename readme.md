# 创建新的虚拟环境(我们命名为 pdf_converter):
```
conda activate pdf_converter
```
# 激活虚拟环境:
```
conda activate pdf_converter
```
安装必要的包:
```
# 安装基本依赖
conda install -c conda-forge pywin32
pip install PyPDF2
pip install reportlab
pip install pyinstaller

# 安装 tkinter (如果没有自动安装的话)
conda install tk
```

# 现在可以尝试打包程序:
```
pyinstaller.exe --onefile --windowed word_to_pdf_converter.py
```