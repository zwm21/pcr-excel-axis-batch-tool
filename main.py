import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("轴模板批量处理工具")
        self.root.geometry("700x500")

        self.files = []  # 存储文件路径

        # 顶部标签
        tk.Label(root, text="待处理文件列表：").pack(anchor="w", padx=5, pady=(5, 0))

        # 列表框
        self.listbox = tk.Listbox(root, selectmode=tk.EXTENDED, height=10, width=80)
        self.listbox.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        # 按钮框架
        btn_frame = tk.Frame(root)
        btn_frame.pack(padx=5, pady=5, fill=tk.X)

        add_files_btn = tk.Button(btn_frame, text="添加文件", command=self.add_files)
        add_files_btn.pack(side=tk.LEFT, padx=2)

        add_folder_btn = tk.Button(btn_frame, text="添加文件夹", command=self.add_folder)
        add_folder_btn.pack(side=tk.LEFT, padx=2)

        clear_btn = tk.Button(btn_frame, text="清空列表", command=self.clear_list)
        clear_btn.pack(side=tk.LEFT, padx=2)

        self.start_btn = tk.Button(btn_frame, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side=tk.RIGHT, padx=2)

        # 追加文字选项框架
        opt_frame = tk.Frame(root)
        opt_frame.pack(padx=5, pady=5, fill=tk.X)

        self.append_text_var = tk.BooleanVar(value=True)
        append_cb = tk.Checkbutton(opt_frame, text="在轴标题末尾追加文字", variable=self.append_text_var)
        append_cb.pack(side=tk.LEFT, padx=2)

        tk.Label(opt_frame, text="文字内容：").pack(side=tk.LEFT, padx=(10,2))
        self.text_suffix_var = tk.StringVar(value=" by 筱娅")
        suffix_entry = tk.Entry(opt_frame, textvariable=self.text_suffix_var, width=15)
        suffix_entry.pack(side=tk.LEFT)

        # 日志区
        tk.Label(root, text="处理日志：").pack(anchor="w", padx=5)
        self.log_text = scrolledtext.ScrolledText(root, width=80, height=12, state=tk.DISABLED)
        self.log_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

    def add_files(self):
        """添加选中的excel文件"""
        paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, p)

    def add_folder(self):
        """添加文件夹中的所有excel文件"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return
        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".xlsx"):
                    full = os.path.join(root, f)
                    if full not in self.files:
                        self.files.append(full)
                        self.listbox.insert(tk.END, full)

    def clear_list(self):
        """清空文件列表"""
        self.files.clear()
        self.listbox.delete(0, tk.END)

    def log(self, msg):
        """向日志区域添加一条消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def start_processing(self):
        """启动处理线程"""
        if not self.files:
            messagebox.showwarning("警告", "没有待处理的文件，请先添加文件或文件夹。")
            return
        self.start_btn.config(state=tk.DISABLED)
        self.log("开始处理...")
        thread = threading.Thread(target=self.process_all, daemon=True)
        thread.start()

    def process_all(self):
        """依次处理每个文件"""
        success_count = 0
        fail_count = 0
        for idx, filepath in enumerate(self.files, 1):
            self.log(f"[{idx}/{len(self.files)}] 正在处理: {os.path.basename(filepath)}")
            try:
                out_path = self.process_file(filepath)
                self.log(f"✓ 处理成功 -> {out_path}")
                success_count += 1
            except Exception as e:
                self.log(f"✗ 处理失败: {str(e)}")
                fail_count += 1
        self.log(f"处理完成！成功: {success_count}, 失败: {fail_count}")
        self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))

    def process_file(self, filepath):
        """处理单个文件：查找表头、合并表头和数据行、保存新文件"""
        wb = load_workbook(filepath)
        if "轴模板" not in wb.sheetnames:
            raise ValueError("找不到工作表 '轴模板'")
        ws = wb["轴模板"]

        # 查找表头位置（返回秒数所在行和列）
        pos = self.find_header(ws, max_row=30)
        if pos is None:
            raise ValueError("前30行内未找到符合要求的表头（秒数|角色|操作|操作|操作|伤害）")
        header_row, start_col = pos   # start_col 是“秒数”所在列
        role_col = start_col + 1      # “角色”所在列
        action_col = start_col + 2    # 第一个“操作”所在列
        # 合并区域：从角色列到伤害列，共5列
        merge_start_col = role_col
        merge_end_col = start_col + 5   # 伤害列

        # 1. 处理表头行
        self.unmerge_in_rect(ws, header_row, merge_start_col, header_row, merge_end_col)
        merge_range = f"{get_column_letter(merge_start_col)}{header_row}:{get_column_letter(merge_end_col)}{header_row}"
        ws.merge_cells(merge_range)
        header_cell = ws.cell(row=header_row, column=merge_start_col)
        header_cell.value = "角色"
        header_cell.alignment = Alignment(horizontal='center', vertical='center')

        # 2. 处理数据行（从表头的下一行开始）
        current_row = header_row + 1
        while True:
            # 检查当前行的角色列是否为空（表示数据结束）
            role_value = ws.cell(row=current_row, column=role_col).value
            if role_value is None or (isinstance(role_value, str) and role_value.strip() == ""):
                break
            
            # 在处理数据行循环内部，读取 role_value 之后，添加以下跳过逻辑：
            merged = None
            for m in ws.merged_cells.ranges:
                if m.min_row <= current_row <= m.max_row and m.min_col <= role_col <= m.max_col:
                    merged = m
                    break
            if merged and (merged.max_col - merged.min_col + 1) > 5:
                current_row += 1
                continue

            # 读取操作文本（第一个操作列）
            op_value = ws.cell(row=current_row, column=action_col).value
            if op_value is None:
                op_text = ""
            else:
                op_text = str(op_value).strip()

            # 决定填充文本和字体颜色
            if op_text == "" or op_text == "连点":
                fill_text = str(role_value).strip()
                font_color = None
            elif op_text == "AUTO":
                fill_text = f"{role_value}(AUTO)"
                font_color = "0000FF"  # 蓝色
            else:
                fill_text = f"{role_value}({op_text})"
                font_color = "FF0000"  # 红色

            # 解除该行的合并区域
            self.unmerge_in_rect(ws, current_row, merge_start_col, current_row, merge_end_col)
            # 合并5格
            merge_range_data = f"{get_column_letter(merge_start_col)}{current_row}:{get_column_letter(merge_end_col)}{current_row}"
            ws.merge_cells(merge_range_data)
            data_cell = ws.cell(row=current_row, column=merge_start_col)
            data_cell.value = fill_text
            data_cell.alignment = Alignment(horizontal='center', vertical='center')
            if font_color:
                data_cell.font = Font(color=font_color)

            current_row += 1

        # --- 新增功能：字体、去粗、追加文本 ---
        # 1. 设置所有单元格字体为“汉仪文黑-65W”（如果系统不支持会自动回退）
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.font = Font(name='汉仪文黑-65W')

        # 2. 取消 B1 和 C1 的加粗
        b1 = ws['B1']
        c1 = ws['C1']
        if b1.font and b1.font.bold:
            b1.font = Font(name='汉仪文黑-65W', bold=False)
        if c1.font and c1.font.bold:
            c1.font = Font(name='汉仪文黑-65W', bold=False)

        # 3. 若启用追加文字，则在 C1 末尾添加用户自定义内容
        if self.append_text_var.get():
            suffix = self.text_suffix_var.get().strip()
            if suffix:
                old_val = str(c1.value) if c1.value else ''
                c1.value = old_val + suffix

        # 生成新文件路径
        dirname = os.path.dirname(filepath)
        basename = os.path.basename(filepath)
        new_name = "已抄轴_" + basename
        new_path = os.path.join(dirname, new_name)

        # 保存
        wb.save(new_path)
        return new_path

    def find_header(self, ws, max_row=30):
        """在指定行数内查找表头，返回(行号, 秒数所在列号)"""
        for row in range(1, max_row + 1):
            for col in range(1, ws.max_column - 5):
                values = []
                for offset in range(6):
                    cell = ws.cell(row=row, column=col + offset)
                    val = cell.value
                    if isinstance(val, str):
                        val = val.strip()
                    else:
                        val = str(val).strip() if val is not None else ""
                    values.append(val)
                if self.match_pattern(values):
                    return row, col
        return None

    @staticmethod
    def match_pattern(vals):
        """检查6个值是否符合表头模式"""
        if len(vals) != 6:
            return False
        # 第1个必须是"秒数"
        if vals[0] != "秒数":
            return False
        # 第2个必须是"角色"
        if vals[1] != "角色":
            return False
        # 第3个必须是"操作"
        if vals[2] != "操作":
            return False
        # 第6个必须是"伤害"
        if vals[5] != "伤害":
            return False
        # 第4和第5个可以是"操作"或空字符串（合并导致）
        ok4 = vals[3] in ("操作", "")
        ok5 = vals[4] in ("操作", "")
        return ok4 and ok5

    @staticmethod
    def unmerge_in_rect(ws, min_row, min_col, max_row, max_col):
        """解除与指定矩形区域相交的所有合并单元格"""
        to_unmerge = []
        for merged in ws.merged_cells.ranges:
            # 检查是否有交集
            if (merged.min_row <= max_row and merged.max_row >= min_row and
                merged.min_col <= max_col and merged.max_col >= min_col):
                to_unmerge.append(merged)
        for rng in to_unmerge:
            ws.unmerge_cells(str(rng))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()