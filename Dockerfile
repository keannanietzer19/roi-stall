# =============================================
# 保本ROI计算器增强版 - Docker 打包配置
# 支持 Mac / Linux / Windows 主机，一键生成 Windows .exe
# =============================================

# 使用专门的 Windows 交叉编译镜像（cdgranier 维护，稳定好用）
FROM cdgranier/pyinstaller-windows:python3-64bit

# 设置工作目录
WORKDIR /src

# 复制代码 + 专属屎哥图标
COPY 保本ROI计算器_增强版.py .
COPY 屎哥.ico .

# 安装额外依赖（ttkbootstrap / matplotlib / openpyxl）
RUN pip install ttkbootstrap matplotlib openpyxl --quiet

# 执行打包（带专属屎哥图标）
RUN pyinstaller \
    --onefile \
    --windowed \
    --icon=屎哥.ico \
    --name "保本ROI计算器增强版_屎哥版" \
    保本ROI计算器_增强版.py

# 打包完成后，exe 文件位于 /src/dist/ 目录
# 你可以把这个目录映射出来下载

# 使用说明（见下方命令）
