try:
    from wcwidth import wcswidth
    HAS_WCWIDTH = True
except ImportError:
    HAS_WCWIDTH = False
    # 简单回退：粗略估算中文字符宽度为2
    def simple_width(s):
        width = 0
        for ch in s:
            if '\u4e00' <= ch <= '\u9fff':
                width += 2
            else:
                width += 1
        return width
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


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
        wb = load_workbook(filepath, rich_text=True)
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

            # 创建富文本对象
            rt = CellRichText()
            if op_text == "" or op_text == "连点":
                fill_text = str(role_value).strip()
                # 黑色文字也显式指定字体，确保是汉仪文黑-65W
                rt.append(TextBlock(InlineFont(rFont='汉仪文黑-65W'), fill_text))
            elif op_text == "AUTO":
                fill_text = f"{role_value}(AUTO)"
                font = InlineFont(rFont='汉仪文黑-65W', color='00b0f0')
                rt.append(TextBlock(font, fill_text))
            else:
                fill_text = f"{role_value}({op_text})"
                font = InlineFont(rFont='汉仪文黑-65W', color='FF0000')
                rt.append(TextBlock(font, fill_text))

            # 解除该行的合并区域
            self.unmerge_in_rect(ws, current_row, merge_start_col, current_row, merge_end_col)
            # 合并5格
            merge_range_data = f"{get_column_letter(merge_start_col)}{current_row}:{get_column_letter(merge_end_col)}{current_row}"
            ws.merge_cells(merge_range_data)
            data_cell = ws.cell(row=current_row, column=merge_start_col)
            data_cell.value = rt
            data_cell.alignment = Alignment(horizontal='center', vertical='center')
            # 不需要再单独设置字体，因为富文本已经包含了字体颜色信息

            current_row += 1

        # --- 新增功能：字体、去粗、追加文本 ---
        # --- 修改字体：仅改变字体名称，保留其他属性 ---
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    old_font = cell.font
                    if old_font is not None:
                        # 复制原有字体的所有属性，仅修改 name
                        cell.font = Font(
                            name='汉仪文黑-65W',
                            size=old_font.size,
                            bold=old_font.bold,
                            italic=old_font.italic,
                            underline=old_font.underline,
                            strike=old_font.strike,
                            color=old_font.color,
                            scheme=old_font.scheme,
                            family=old_font.family,
                            charset=old_font.charset
                        )
                    else:
                        # 没有原字体则只设置名称（大小等采用默认，通常为11）
                        cell.font = Font(name='汉仪文黑-65W')

        # 2. 取消 B1 和 C1 的加粗（同时保留其他属性）
        for cell in [ws['B1'], ws['C1']]:
            if cell.value is not None:
                old_font = cell.font
                if old_font is not None:
                    cell.font = Font(
                        name='汉仪文黑-65W',
                        size=old_font.size,
                        bold=False,               # 取消加粗
                        italic=old_font.italic,
                        underline=old_font.underline,
                        strike=old_font.strike,
                        color=old_font.color,
                        scheme=old_font.scheme,
                        family=old_font.family,
                        charset=old_font.charset
                    )
                else:
                    cell.font = Font(name='汉仪文黑-65W', bold=False)

        # 3. 若启用追加文字，则在 C1 末尾添加用户自定义内容
        c1 = ws['C1']
        if self.append_text_var.get():
            suffix = self.text_suffix_var.get()
            if suffix:
                old_val = str(c1.value) if c1.value else ''
                c1.value = old_val + suffix

        # ========== 合并符合条件的相邻行（完整规则） ==========
        # 收集所有有内容的数据行号
        rows = []
        current = header_row + 1
        while current <= ws.max_row:
            cell = ws.cell(row=current, column=merge_start_col)
            if cell.value is not None:
                rows.append(current)
            current += 1

        # 辅助函数：判断某行是否属于宽度超过5的合并单元格
        def is_oversize_merged(row):
            for m in ws.merged_cells.ranges:
                if m.min_row <= row <= m.max_row and m.min_col <= merge_start_col <= m.max_col:
                    if m.max_col - m.min_col + 1 > 5:
                        return True
            return False

        to_delete = []
        i = 0
        while i < len(rows) - 1:
            cur_row = rows[i]
            nxt_row = rows[i+1]

            # 最后一行不参与合并
            if nxt_row == rows[-1]:
                i += 1
                continue
            # 当前行或下一行的角色格宽度超过5，跳过
            if is_oversize_merged(cur_row) or is_oversize_merged(nxt_row):
                i += 1
                continue

            cur_cell = ws.cell(row=cur_row, column=merge_start_col)
            nxt_cell = ws.cell(row=nxt_row, column=merge_start_col)

            # 读取秒数
            cur_sec = str(ws.cell(row=cur_row, column=start_col).value or "").strip()
            nxt_sec = str(ws.cell(row=nxt_row, column=start_col).value or "").strip()

            # 提取纯文本
            def get_plain_text(value):
                if isinstance(value, CellRichText):
                    return "".join(str(b) for b in value)
                return str(value) if value else ""

            cur_text = get_plain_text(cur_cell.value)
            nxt_text = get_plain_text(nxt_cell.value)

            # 模拟合并后文本，用于宽度判断
            if cur_sec == nxt_sec:
                merged_text = cur_text + "-" + nxt_text
            else:
                if "(AUTO)" in nxt_text:
                    new_nxt = nxt_text.replace("(AUTO)", f"({nxt_sec} AUTO)")
                elif "(" in nxt_text and ")" in nxt_text:
                    start = nxt_text.find("(")
                    end = nxt_text.find(")")
                    if start != -1 and end != -1:
                        inner = nxt_text[start+1:end]
                        new_nxt = nxt_text[:start+1] + nxt_sec + " " + inner + nxt_text[end:]
                    else:
                        new_nxt = nxt_text + f"({nxt_sec})"
                else:
                    new_nxt = nxt_text + f"({nxt_sec})"
                merged_text = cur_text + "-" + new_nxt

            if self.display_width(merged_text) > 31:
                i += 1
                continue

            # ---------- 开始构建新富文本 ----------
            new_rt = CellRichText()

            # 辅助函数：确保块带有目标字体
            def ensure_font(block):
                if isinstance(block, TextBlock):
                    if block.font is None or block.font.rFont is None:
                        old_color = block.font.color if block.font else None
                        new_font = InlineFont(rFont='汉仪文黑-65W', color=old_color)
                        return TextBlock(new_font, block.text)
                    else:
                        return block
                else:
                    return TextBlock(InlineFont(rFont='汉仪文黑-65W'), str(block))

            # 1) 添加当前行所有内容（统一字体）
            if isinstance(cur_cell.value, CellRichText):
                for block in cur_cell.value:
                    new_rt.append(ensure_font(block))
            else:
                new_rt.append(ensure_font(str(cur_cell.value) if cur_cell.value else ""))

            # 2) 添加分隔符 "-"，带字体
            new_rt.append(TextBlock(InlineFont(rFont='汉仪文黑-65W'), "-"))

            # 3) 处理下一行（保留颜色，统一字体）
            if isinstance(nxt_cell.value, CellRichText):
                nxt_blocks = list(nxt_cell.value)
                if cur_sec != nxt_sec:
                    if len(nxt_blocks) == 1 and isinstance(nxt_blocks[0], TextBlock):
                        original_color = nxt_blocks[0].font.color if nxt_blocks[0].font else None
                        # 生成修改后的文本
                        if "(AUTO)" in nxt_text:
                            modified = nxt_text.replace("(AUTO)", f"({nxt_sec} AUTO)")
                        elif "(" in nxt_text and ")" in nxt_text:
                            start = nxt_text.find("(")
                            end = nxt_text.find(")")
                            inner = nxt_text[start+1:end] if start+1 < end else ""
                            modified = nxt_text[:start+1] + nxt_sec + " " + inner + nxt_text[end:]
                        else:
                            modified = nxt_text + f"({nxt_sec})"
                        # 创建新块，保留原颜色并强制使用汉仪文黑
                        new_rt.append(TextBlock(InlineFont(rFont='汉仪文黑-65W', color=original_color), modified))
                    else:
                        # 多块或混合，先原样添加（但统一字体）
                        for block in nxt_blocks:
                            new_rt.append(ensure_font(block))
                        # 追加纯文本秒数
                        new_rt.append(TextBlock(InlineFont(rFont='汉仪文黑-65W'), f"({nxt_sec})"))
                else:
                    # 秒数相同，直接添加所有块（统一字体）
                    for block in nxt_blocks:
                        new_rt.append(ensure_font(block))
            else:
                # 下一行是普通字符串，处理并带字体
                nxt_str = str(nxt_cell.value) if nxt_cell.value else ""
                if cur_sec != nxt_sec:
                    if "(AUTO)" in nxt_str:
                        nxt_str = nxt_str.replace("(AUTO)", f"({nxt_sec} AUTO)")
                    elif "(" in nxt_str and ")" in nxt_str:
                        start = nxt_str.find("(")
                        end = nxt_str.find(")")
                        inner = nxt_str[start+1:end]
                        nxt_str = nxt_str[:start+1] + nxt_sec + " " + inner + nxt_str[end:]
                    else:
                        nxt_str = nxt_str + f"({nxt_sec})"
                new_rt.append(TextBlock(InlineFont(rFont='汉仪文黑-65W'), nxt_str))

            # 写入合并结果
            cur_cell.value = new_rt
            cur_cell.alignment = Alignment(horizontal='center', vertical='center')

            to_delete.append(nxt_row)
            rows.pop(i + 1)

        # 删除被合并的行
        #for del_row in sorted(to_delete, reverse=True):
        #    if 1 <= del_row <= ws.max_row:
        #        ws.delete_rows(del_row)

        # 不删除，而是将下一行内容清空并隐藏
        for del_row in sorted(to_delete, reverse=True):
            if 1 <= del_row <= ws.max_row:
                ws.row_dimensions[del_row].hidden = True
                # 清空合并单元格的内容，但保留合并格式（不解除）
                cell = ws.cell(row=del_row, column=merge_start_col)
                cell.value = None

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
    
    @staticmethod
    def display_width(text):
        if HAS_WCWIDTH:
            return wcswidth(text)
        else:
            return simple_width(text)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()