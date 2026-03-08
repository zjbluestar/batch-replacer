# -*- coding: utf-8 -*-
"""
批量关键词替换工具 v1.0
功能：批量修改Word/Excel文件中的关键词
支持格式：.doc, .docx, .xls, .xlsx
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# 尝试导入文档处理库
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


class BatchReplacerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量关键词替换工具 v1.0")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 替换规则列表
        self.replace_rules = []
        
        self.create_widgets()
        self.check_dependencies()
    
    def check_dependencies(self):
        """检查依赖库"""
        missing = []
        if not HAS_DOCX:
            missing.append("python-docx (处理.docx)")
        if not HAS_OPENPYXL:
            missing.append("openpyxl (处理.xlsx)")
        if not HAS_WIN32COM:
            missing.append("pywin32 (处理.doc/.xls)")
        
        if missing:
            self.log(f"⚠️ 缺少依赖库: {', '.join(missing)}")
            self.log("请安装: pip install python-docx openpyxl pywin32")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件夹选择
        folder_frame = ttk.LabelFrame(main_frame, text="选择文件夹", padding="5")
        folder_frame.pack(fill=tk.X, pady=5)
        
        self.folder_path = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.folder_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_frame, text="浏览...", command=self.select_folder).pack(side=tk.LEFT)
        
        # 替换规则
        rules_frame = ttk.LabelFrame(main_frame, text="替换规则", padding="5")
        rules_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 规则输入
        input_frame = ttk.Frame(rules_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="旧词:").pack(side=tk.LEFT, padx=5)
        self.old_word = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.old_word, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(input_frame, text="→").pack(side=tk.LEFT)
        
        ttk.Label(input_frame, text="新词:").pack(side=tk.LEFT, padx=5)
        self.new_word = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.new_word, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(input_frame, text="添加规则", command=self.add_rule).pack(side=tk.LEFT, padx=10)
        
        # 规则列表
        list_frame = ttk.Frame(rules_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("old", "new")
        self.rules_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=5)
        self.rules_tree.heading("old", text="旧词")
        self.rules_tree.heading("new", text="新词")
        self.rules_tree.column("old", width=200)
        self.rules_tree.column("new", width=200)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=scrollbar.set)
        
        self.rules_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 删除规则按钮
        btn_frame = ttk.Frame(rules_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="删除选中规则", command=self.delete_rule).pack(side=tk.RIGHT)
        
        # 输出选项
        output_frame = ttk.LabelFrame(main_frame, text="输出选项", padding="5")
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_option = tk.StringVar(value="new_folder")
        ttk.Radiobutton(output_frame, text="输出到新文件夹", variable=self.output_option, value="new_folder").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(output_frame, text="直接覆盖原文件", variable=self.output_option, value="overwrite").pack(side=tk.LEFT, padx=10)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=8, state=tk.DISABLED)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="开始替换", command=self.start_replace).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.RIGHT, padx=10)
    
    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含Word/Excel文件的文件夹")
        if folder:
            self.folder_path.set(folder)
            self.log(f"已选择文件夹: {folder}")
    
    def add_rule(self):
        """添加替换规则"""
        old = self.old_word.get().strip()
        new = self.new_word.get().strip()
        
        if not old:
            messagebox.showwarning("警告", "请输入旧词")
            return
        
        self.replace_rules.append((old, new))
        self.rules_tree.insert("", tk.END, values=(old, new))
        self.log(f"添加规则: '{old}' → '{new}'")
        
        # 清空输入
        self.old_word.set("")
        self.new_word.set("")
    
    def delete_rule(self):
        """删除选中的规则"""
        selected = self.rules_tree.selection()
        if not selected:
            return
        
        for item in selected:
            index = self.rules_tree.index(item)
            old, new = self.replace_rules[index]
            self.replace_rules.pop(index)
            self.rules_tree.delete(item)
            self.log(f"删除规则: '{old}' → '{new}'")
    
    def log(self, message):
        """添加日志"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
    
    def start_replace(self):
        """开始替换"""
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("错误", "请先选择文件夹")
            return
        
        if not self.replace_rules:
            messagebox.showerror("错误", "请添加至少一条替换规则")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("错误", "选择的文件夹不存在")
            return
        
        # 创建输出文件夹
        if self.output_option.get() == "new_folder":
            output_folder = os.path.join(folder, "替换后文件")
            os.makedirs(output_folder, exist_ok=True)
        else:
            output_folder = folder
        
        self.log("=" * 50)
        self.log("开始处理...")
        
        # 支持的文件扩展名
        extensions = ['.doc', '.docx', '.xls', '.xlsx']
        
        processed = 0
        errors = 0
        
        for root, dirs, files in os.walk(folder):
            # 跳过输出文件夹
            if "替换后文件" in root:
                continue
            
            for filename in files:
                ext = os.path.splitext(filename)[1].lower()
                if ext in extensions:
                    filepath = os.path.join(root, filename)
                    
                    try:
                        if ext in ['.docx']:
                            self.replace_docx(filepath, output_folder)
                        elif ext in ['.doc']:
                            self.replace_doc(filepath, output_folder)
                        elif ext in ['.xlsx']:
                            self.replace_xlsx(filepath, output_folder)
                        elif ext in ['.xls']:
                            self.replace_xls(filepath, output_folder)
                        
                        processed += 1
                        self.log(f"✓ {filename}")
                    except Exception as e:
                        errors += 1
                        self.log(f"✗ {filename}: {str(e)}")
        
        self.log("=" * 50)
        self.log(f"处理完成！成功: {processed}, 失败: {errors}")
        messagebox.showinfo("完成", f"处理完成！\n成功: {processed}\n失败: {errors}")
    
    def replace_docx(self, filepath, output_folder):
        """替换.docx文件中的关键词"""
        if not HAS_DOCX:
            raise Exception("缺少python-docx库")
        
        doc = Document(filepath)
        
        for old, new in self.replace_rules:
            # 替换段落中的文本
            for paragraph in doc.paragraphs:
                if old in paragraph.text:
                    # 需要保留格式，逐个run处理
                    for run in paragraph.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)
            
            # 替换表格中的文本
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if old in cell.text:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    if old in run.text:
                                        run.text = run.text.replace(old, new)
        
        # 保存文件
        filename = os.path.basename(filepath)
        output_path = os.path.join(output_folder, filename)
        doc.save(output_path)
    
    def replace_doc(self, filepath, output_folder):
        """替换.doc文件中的关键词（使用Word COM）"""
        if not HAS_WIN32COM:
            raise Exception("缺少pywin32库，无法处理.doc文件")
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            doc = word.Documents.Open(filepath)
            
            for old, new in self.replace_rules:
                # 使用Word的查找替换功能
                find = doc.Content.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Execute(old, False, False, False, False, False, True, 1, True, new, 2)
            
            filename = os.path.basename(filepath)
            output_path = os.path.join(output_folder, filename)
            doc.SaveAs(output_path)
            doc.Close()
        finally:
            word.Quit()
    
    def replace_xlsx(self, filepath, output_folder):
        """替换.xlsx文件中的关键词"""
        if not HAS_OPENPYXL:
            raise Exception("缺少openpyxl库")
        
        wb = load_workbook(filepath)
        
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        for old, new in self.replace_rules:
                            if old in cell.value:
                                cell.value = cell.value.replace(old, new)
        
        filename = os.path.basename(filepath)
        output_path = os.path.join(output_folder, filename)
        wb.save(output_path)
    
    def replace_xls(self, filepath, output_folder):
        """替换.xls文件中的关键词（使用Excel COM）"""
        if not HAS_WIN32COM:
            raise Exception("缺少pywin32库，无法处理.xls文件")
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            wb = excel.Workbooks.Open(filepath)
            
            for sheet in wb.Sheets:
                used_range = sheet.UsedRange
                for row in used_range.Rows:
                    for cell in row.Cells:
                        if cell.Value and isinstance(cell.Value, str):
                            for old, new in self.replace_rules:
                                if old in cell.Value:
                                    cell.Value = cell.Value.replace(old, new)
            
            filename = os.path.basename(filepath)
            output_path = os.path.join(output_folder, filename)
            wb.SaveAs(output_path)
            wb.Close()
        finally:
            excel.Quit()


def main():
    root = tk.Tk()
    app = BatchReplacerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
