# 批量关键词替换工具 v1.0

## 功能

批量修改 Word/Excel 文件中的关键词，支持：
- Word: .doc, .docx
- Excel: .xls, .xlsx

## 使用方法

### 方式一：Python环境运行

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行程序：
```bash
python batch_replacer.py
```

### 方式二：打包成exe（无需Python）

在一台有Python环境的电脑上：

1. 安装打包工具：
```bash
pip install pyinstaller
```

2. 打包命令：
```bash
pyinstaller --onefile --windowed --name "批量关键词替换工具" batch_replacer.py
```

3. 打包完成后，在 `dist` 文件夹中找到 `批量关键词替换工具.exe`

4. 将exe文件拷贝到办公电脑使用

## 注意事项

- 处理.doc和.xls文件需要电脑上安装Microsoft Office
- .docx和.xlsx文件不需要安装Office
- 建议先备份文件再批量替换

## 界面说明

1. **选择文件夹**：选择包含要处理的 Word/Excel 文件
2. **添加规则**：输入旧词、新词，点击"添加规则"
3. **输出选项**：选择输出到新文件夹或覆盖原文件
4. **开始替换**：点击开始处理
# Build test 2026年 3月 8日 星期日 22时20分27秒 CST
