# 保本ROI计算器增强版 - Docker 打包指南（Mac 用户）

## 前提
- 已安装 Docker Desktop（Mac）
- 已下载 `保本ROI计算器_增强版.py` 和 `Dockerfile` 到同一文件夹

## 方法一：使用 Dockerfile 构建（推荐）

```bash
# 1. 进入存放 py 文件和 Dockerfile 的文件夹
cd /path/to/your/folder

# 2. 构建镜像（第一次会下载 ~1.5GB 镜像，需耐心等待）
docker build -t roi-屎哥版 .

# 3. 运行容器并把 dist 文件夹复制出来
docker run --rm -v "$(pwd)/dist:/output" roi-屎哥版 cp -r /src/dist/. /output/

# 4. 完成！现在 dist 文件夹里就有 Windows exe 了
open dist
```

## 方法二：直接用现成镜像（更快，不用写 Dockerfile）

```bash
# 在你的代码文件夹下执行下面命令即可

docker run --rm \
  -v "$(pwd):/src/" \
  cdgranier/pyinstaller-windows:python3-64bit \
  pyinstaller \
    --onefile \
    --windowed \
    --name "保本ROI计算器增强版_屎哥版" \
    /src/保本ROI计算器_增强版.py

# 打包完成后直接查看
open dist
```

## 注意事项
- 第一次运行会自动下载镜像（约 1.2~1.8GB），请保持网络通畅
- 生成的 exe 是 **64位 Windows** 版本，可直接在 Windows 上运行
- 如果提示权限问题，Mac 上可能需要 `sudo`（不推荐）或在 Docker Desktop 设置里允许文件共享

## 生成的文件
- `保本ROI计算器增强版_屎哥版.exe` （带红色「屎哥」 + 毛利润核心指标）

有问题随时问我！
