# 打包指南

## GitHub Actions 自动打包（推荐）

### 步骤

1. 在GitHub创建新仓库
2. 上传整个 `batch-replacer` 文件夹
3. GitHub Actions自动构建
4. 下载生成的exe文件

### 详细命令

```bash
# 1. 创建GitHub仓库（在GitHub网页操作）
# 仓库名：batch-replacer

# 2. 初始化git并推送
cd /Users/jintianyeyaozhuanqianya/Documents/00_OpenClaw_Inbox/tools/batch-replacer
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/zjbluestar/batch-replacer.git
git push -u origin main

# 3. 等待GitHub Actions构建完成
# 4. 在Actions页面下载exe文件
```

## 手动打包（需要Windows电脑）

如果需要在Windows电脑上手动打包：

```bash
# 安装依赖
pip install pyinstaller python-docx openpyxl pywin32

# 打包
pyinstaller --onefile --windowed --name "批量关键词替换工具" batch_replacer.py
```

打包后的exe在 `dist` 文件夹中。
