# 批量关键词替换工具 - Windows打包详细步骤

> 本文档教你如何在Windows电脑上将Python脚本打包成exe文件

---

## 📋 准备工作

### 需要的东西
- ✅ 一台Windows电脑（Win10/Win11）
- ✅ 管理员权限（用于安装软件）
- ✅ 约30分钟时间

### 文件下载
从以下地址下载代码，或者让阿健把 `batch-replacer` 文件夹发给你：
- GitHub: https://github.com/zjbluestar/batch-replacer

---

## 第一步：安装Python

### 1.1 下载Python

1. 打开浏览器，访问：https://www.python.org/downloads/
2. 点击 **Download Python 3.11.x** 按钮
3. 等待下载完成（约25MB）

### 1.2 安装Python

⚠️ **非常重要！**

1. 双击下载的安装文件
2. **必须勾选** 底部的 `Add Python to PATH` ✅
3. 点击 `Install Now`
4. 等待安装完成
5. 点击 `Close`

### 1.3 验证安装

1. 按 `Win + R`，输入 `cmd`，回车
2. 在命令提示符中输入：
   ```cmd
   python --version
   ```
3. 应该显示：`Python 3.11.x`

---

## 第二步：准备代码文件

### 2.1 创建文件夹

1. 在桌面创建文件夹 `batch-replacer`
2. 把以下文件放入文件夹：
   - `batch_replacer.py`
   - `requirements.txt`

### 2.2 确认文件结构

```
桌面\batch-replacer\
├── batch_replacer.py    ← 主程序
└── requirements.txt     ← 依赖列表
```

---

## 第三步：安装依赖

### 3.1 打开命令提示符

1. 按 `Win + R`
2. 输入 `cmd`
3. 回车

### 3.2 进入文件夹

输入以下命令（假设在桌面）：

```cmd
cd %USERPROFILE%\Desktop\batch-replacer
```

### 3.3 安装依赖库

```cmd
pip install -r requirements.txt
pip install pyinstaller
```

等待安装完成（约1-2分钟）

**预期输出：**
```
Successfully installed python-docx-xxx openpyxl-xxx pywin32-xxx pyinstaller-xxx
```

---

## 第四步：测试脚本

### 4.1 运行测试

```cmd
python batch_replacer.py
```

**预期结果：** 弹出一个窗口界面

### 4.2 如果弹出窗口

- 说明脚本正常工作
- 关闭窗口，继续下一步打包

### 4.3 如果报错

常见问题：
- `No module named 'docx'` → 依赖没装好，重试第三步
- `'pip' 不是内部命令` → Python没加到PATH，重装Python

---

## 第五步：打包成exe

### 5.1 执行打包命令

```cmd
pyinstaller --onefile --windowed --name "批量关键词替换工具" batch_replacer.py
```

**参数说明：**
- `--onefile`：打包成单个exe文件
- `--windowed`：运行时不显示黑色命令窗口
- `--name "xxx"`：指定exe文件名

### 5.2 等待打包

打包过程约1-3分钟，输出类似：

```
Building EXE from EXE-00.toc completed successfully.
```

---

## 第六步：获取exe文件

### 6.1 找到exe

打包完成后，文件结构：

```
batch-replacer\
├── dist\
│   └── 批量关键词替换工具.exe  ← 这是你要的！
├── build\
├── batch_replacer.spec
└── ...
```

### 6.2 复制exe

1. 进入 `dist` 文件夹
2. 找到 `批量关键词替换工具.exe`
3. 复制到你想用的地方（如U盘）

---

## 第七步：使用exe

### 7.1 无需安装

将 `批量关键词替换工具.exe` 复制到任何Windows电脑，双击即可运行！

### 7.2 首次运行提示

如果系统提示"Windows已保护你的电脑"：
1. 点击"更多信息"
2. 点击"仍要运行"

---

## 🎯 完整命令汇总

```cmd
# 1. 进入文件夹
cd %USERPROFILE%\Desktop\batch-replacer

# 2. 安装依赖
pip install -r requirements.txt
pip install pyinstaller

# 3. 测试运行
python batch_replacer.py

# 4. 打包
pyinstaller --onefile --windowed --name "批量关键词替换工具" batch_replacer.py

# 5. exe在 dist 文件夹中
```

---

## ❓ 常见问题

### Q1: pip安装很慢/失败
使用国内镜像：
```cmd
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### Q2: 打包后exe太大
正常！包含了Python运行环境，约50-80MB。

### Q3: exe在其他电脑运行报错
确保目标电脑：
- 是Windows系统（Win7及以上）
- 安装了Microsoft Office（处理.doc/.xls时需要）

### Q4: 只需要处理.docx/.xlsx
如果不需要处理旧格式文件，可以：
```cmd
pip install python-docx openpyxl pyinstaller
```
（不需要 pywin32）

### Q5: 打包时出现编码错误
在打包命令前添加：
```cmd
chcp 65001
pyinstaller --onefile --windowed --name "批量关键词替换工具" batch_replacer.py
```

---

## 📸 界面预览

```
┌─────────────────────────────────────────────┐
│   批量关键词替换工具 v1.0                    │
├─────────────────────────────────────────────┤
│ 选择文件夹: [C:\Users\用户\Desktop\文件]     │
│                                             │
│ 替换规则:                                   │
│   旧词: [公司A      ] → 新词: [公司B      ] │
│   [+ 添加规则]                              │
│                                             │
│   ┌──────────────┬──────────────┐          │
│   │     旧词     │     新词     │          │
│   ├──────────────┼──────────────┤          │
│   │    公司A     │    公司B     │          │
│   │  北京朝阳区   │  上海浦东    │          │
│   │    张三      │    李四      │          │
│   └──────────────┴──────────────┘          │
│                                             │
│ 输出: ☑ 输出到新文件夹                       │
│                                             │
│ [开始替换]                        [退出]    │
├─────────────────────────────────────────────┤
│ 处理日志:                                   │
│ ✓ 合同模板.docx                             │
│ ✓ 客户名单.xlsx                             │
│ 完成！成功: 2, 失败: 0                      │
└─────────────────────────────────────────────┘
```

---

## 📞 技术支持

如有问题，联系：阿健

文档版本：v1.0
更新日期：2026-03-08
