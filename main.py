import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from ui_config import *

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# 模板系统
import templates as tmpl


# ============================================================
#  样式工具函数
# ============================================================
def _make_font(size=FONT_SIZE_NORMAL, bold=False, family=FONT_FAMILY):
    weight = "bold" if bold else "normal"
    return (family, size, weight)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("轴模板批量处理工具")
        self.root.geometry(f"{WIN_DEFAULT_WIDTH}x{WIN_DEFAULT_HEIGHT}")
        self.root.minsize(WIN_MIN_WIDTH, WIN_MIN_HEIGHT)
        self.root.configure(bg=COLOR_BG)

        self.files = []
        self.last_dir = None

        # 模板选择
        self.template_var = tk.StringVar(value=tmpl.DEFAULT_TEMPLATE_ID)

        self._configure_ttk_style()
        self._build_header()                # row=0
        self._build_file_section()          # row=1  stretch
        self._build_template_card()         # row=2
        self._build_option_bar()            # row=3
        self._build_log_section()           # row=4  stretch

        # 缩放权重
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    # ----------------------------------------------------------
    #  ttk 主题
    # ----------------------------------------------------------
    def _configure_ttk_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Section.TLabel",
                        font=_make_font(FONT_SIZE_SMALL, bold=True),
                        foreground=COLOR_TEXT_SECONDARY,
                        background=COLOR_CARD_BG)

        style.configure("Title.TLabel",
                        font=_make_font(FONT_SIZE_HEADING, bold=True),
                        foreground=COLOR_TEXT,
                        background=COLOR_BG)

        style.configure("Primary.TButton",
                        font=_make_font(FONT_SIZE_NORMAL),
                        padding=(PAD_CARD, GAP_MD),
                        background=COLOR_PRIMARY,
                        foreground="#FFFFFF",
                        borderwidth=0, focuscolor="none")
        style.map("Primary.TButton",
                  background=[("active", COLOR_PRIMARY_HOVER),
                              ("pressed", COLOR_PRIMARY_ACTIVE)],
                  foreground=[("active", "#FFFFFF")])

        style.configure("Accent.TButton",
                        font=_make_font(FONT_SIZE_NORMAL, bold=True),
                        padding=(PAD_CARD * 2, GAP_MD),
                        background=COLOR_ACCENT,
                        foreground="#FFFFFF",
                        borderwidth=0, focuscolor="none")
        style.map("Accent.TButton",
                  background=[("active", COLOR_ACCENT_HOVER),
                              ("pressed", COLOR_ACCENT_ACTIVE)],
                  foreground=[("active", "#FFFFFF")])

        style.configure("Danger.TButton",
                        font=_make_font(FONT_SIZE_NORMAL),
                        padding=(PAD_CARD, GAP_MD),
                        background="#E8E8E8",
                        borderwidth=0, focuscolor="none")
        style.map("Danger.TButton",
                  background=[("active", "#D8D8D8"),
                              ("pressed", "#C8C8C8")])

    # ----------------------------------------------------------
    #  顶部标题栏
    # ----------------------------------------------------------
    def _build_header(self):
        frame = tk.Frame(self.root, bg=COLOR_BG)
        frame.grid(row=0, column=0, sticky="ew",
                   padx=PAD_PAGE, pady=(PAD_PAGE, GAP_LG))
        frame.grid_columnconfigure(0, weight=1)
        ttk.Label(frame, text="轴模板批量处理工具",
                  style="Title.TLabel").pack(side=tk.LEFT)

    # ----------------------------------------------------------
    #  文件列表卡片
    # ----------------------------------------------------------
    def _build_file_section(self):
        card = tk.Frame(self.root, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1)
        card.grid(row=1, column=0, sticky="nsew",
                  padx=PAD_PAGE, pady=(0, GAP_LG))
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        title_row = tk.Frame(card, bg=COLOR_CARD_BG)
        title_row.grid(row=0, column=0, sticky="ew",
                       padx=PAD_CARD, pady=(PAD_TITLE_Y, GAP_MD))
        ttk.Label(title_row, text="待处理文件列表",
                  style="Section.TLabel").pack(side=tk.LEFT)
        self._file_count_label = ttk.Label(
            title_row, text="（0 个文件）",
            font=_make_font(FONT_SIZE_SMALL),
            foreground=COLOR_TEXT_LIGHT, background=COLOR_CARD_BG)
        self._file_count_label.pack(side=tk.LEFT, padx=(GAP_MD, 0))

        list_frame = tk.Frame(card, bg=COLOR_CARD_BORDER)
        list_frame.grid(row=1, column=0, sticky="nsew",
                        padx=PAD_CARD, pady=(0, GAP_MD))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED, height=LISTBOX_HEIGHT,
            bg=COLOR_LISTBOX_BG, fg=COLOR_TEXT,
            font=_make_font(FONT_SIZE_NORMAL),
            selectbackground=COLOR_LISTBOX_SELECT,
            selectforeground=COLOR_LISTBOX_SELECT_TEXT,
            activestyle="none", borderwidth=0, highlightthickness=0,
            relief="flat")
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

        btn_row = tk.Frame(card, bg=COLOR_CARD_BG)
        btn_row.grid(row=2, column=0, sticky="ew",
                     padx=PAD_CARD, pady=(0, PAD_CARD_Y))

        self.add_files_btn = ttk.Button(
            btn_row, text="＋ 添加文件", style="Primary.TButton",
            command=self.add_files)
        self.add_files_btn.pack(side=tk.LEFT, padx=(0, GAP_SM))

        self.add_folder_btn = ttk.Button(
            btn_row, text=" 添加文件夹", style="Primary.TButton",
            command=self.add_folder)
        self.add_folder_btn.pack(side=tk.LEFT, padx=GAP_SM)

        self.clear_btn = ttk.Button(
            btn_row, text="清空列表", style="Danger.TButton",
            command=self.clear_list)
        self.clear_btn.pack(side=tk.LEFT, padx=GAP_SM)

    # ----------------------------------------------------------
    #  模板信息卡片（选择器 + 摘要 + 可收起的详细描述）
    # ----------------------------------------------------------
    def _build_template_card(self):
        card = tk.Frame(self.root, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1)
        card.grid(row=2, column=0, sticky="ew",
                  padx=PAD_PAGE, pady=(0, GAP_LG))
        card.grid_columnconfigure(0, weight=1)

        # 标题行 + 选择器 + 显示/隐藏按钮
        sel_row = tk.Frame(card, bg=COLOR_CARD_BG)
        sel_row.grid(row=0, column=0, sticky="ew",
                     padx=PAD_CARD, pady=(PAD_TITLE_Y, GAP_SM))

        ttk.Label(sel_row, text="处理模板",
                  style="Section.TLabel").pack(side=tk.LEFT)

        self.template_combo = ttk.Combobox(
            sel_row, textvariable=self.template_var,
            values=[t["name"] for t in tmpl.TEMPLATES],
            state="readonly", width=28,
            font=_make_font(FONT_SIZE_NORMAL))
        self.template_combo.pack(side=tk.LEFT, padx=(GAP_XL, 0))
        self.template_combo.bind("<<ComboboxSelected>>", self._on_template_changed)

        # 显示/隐藏详细描述按钮
        self._desc_toggle_btn = ttk.Button(
            sel_row, text="▸ 显示详细说明",
            style="Primary.TButton",
            command=self._toggle_desc)
        self._desc_toggle_btn.pack(side=tk.RIGHT)

        # 模板摘要
        self.template_summary_label = tk.Label(
            card, text="", bg=COLOR_CARD_BG, fg=COLOR_TEXT_SECONDARY,
            font=_make_font(FONT_SIZE_SMALL),
            anchor="w", justify=tk.LEFT,
            wraplength=WIN_DEFAULT_WIDTH - 60)
        self.template_summary_label.grid(
            row=1, column=0, sticky="ew",
            padx=PAD_CARD, pady=(0, GAP_SM))

        # 描述文本框（只读，默认隐藏）
        self._desc_visible = False
        self.desc_frame = tk.Frame(card, bg=COLOR_CARD_BORDER)
        self.desc_frame.grid_rowconfigure(0, weight=1)
        self.desc_frame.grid_columnconfigure(0, weight=1)

        self.template_desc_text = tk.Text(
            self.desc_frame,
            height=14, state=tk.DISABLED,
            bg=COLOR_LOG_BG, fg=COLOR_TEXT,
            font=_make_font(FONT_SIZE_SMALL, family=FONT_MONO),
            borderwidth=0, highlightthickness=0,
            relief="flat", wrap=tk.WORD, padx=GAP_MD, pady=GAP_MD)
        self.template_desc_text.grid(row=0, column=0, sticky="nsew",
                                     padx=1, pady=1)
        # 默认隐藏
        self.desc_frame.grid_remove()

        # 初始设置
        self.template_combo.current(0)
        self._update_template_info()

    def _toggle_desc(self):
        """切换详细描述的显示/隐藏。"""
        if self._desc_visible:
            self.desc_frame.grid_remove()
            self._desc_toggle_btn.config(text="▸ 显示详细说明")
            self._desc_visible = False
        else:
            self.desc_frame.grid(row=2, column=0, sticky="ew",
                                padx=PAD_CARD, pady=(0, PAD_CARD_Y))
            self._desc_toggle_btn.config(text="▾ 隐藏详细说明")
            self._desc_visible = True

    def _on_template_changed(self, event=None):
        """模板切换回调。"""
        self._update_template_info()

    def _update_template_info(self):
        """根据当前选中模板更新摘要和描述。"""
        idx = self.template_combo.current()
        if idx < 0:
            idx = 0
        t = tmpl.TEMPLATES[idx]
        self.template_var.set(t["id"])
        self.template_summary_label.config(text=t["summary"])
        # 即便隐藏也预先填充文本，展开时立即可见
        self._set_desc_text(t["description"])

    def _set_desc_text(self, text):
        """向描述文本框写入内容。"""
        self.template_desc_text.config(state=tk.NORMAL)
        self.template_desc_text.delete("1.0", tk.END)
        self.template_desc_text.insert("1.0", text)
        self.template_desc_text.config(state=tk.DISABLED)

    # ----------------------------------------------------------
    #  选项栏 + 开始按钮
    # ----------------------------------------------------------
    def _build_option_bar(self):
        card = tk.Frame(self.root, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1)
        card.grid(row=3, column=0, sticky="ew",
                  padx=PAD_PAGE, pady=(0, GAP_LG))
        card.grid_columnconfigure(1, weight=1)

        # 复选框
        self.append_text_var = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(
            card, text="在轴标题末尾追加文字",
            variable=self.append_text_var,
            bg=COLOR_CARD_BG, fg=COLOR_TEXT,
            font=_make_font(FONT_SIZE_NORMAL),
            activebackground=COLOR_CARD_BG,
            activeforeground=COLOR_TEXT,
            selectcolor=COLOR_CARD_BG)
        cb.grid(row=0, column=0, sticky="w",
                padx=(PAD_CARD, GAP_MD), pady=PAD_CARD_Y)

        # 输入框
        lbl = tk.Label(card, text="内容：", bg=COLOR_CARD_BG,
                       fg=COLOR_TEXT_SECONDARY,
                       font=_make_font(FONT_SIZE_NORMAL))
        lbl.grid(row=0, column=1, sticky="e",
                 padx=(0, GAP_SM), pady=PAD_CARD_Y)

        self.text_suffix_var = tk.StringVar(value=" by 筱娅")
        self.suffix_entry = tk.Entry(
            card, textvariable=self.text_suffix_var,
            width=ENTRY_WIDTH,
            font=_make_font(FONT_SIZE_NORMAL),
            bg="#FFFFFF", fg=COLOR_TEXT,
            insertbackground=COLOR_TEXT,
            relief="solid", borderwidth=1,
            highlightbackground=COLOR_CARD_BORDER,
            highlightthickness=0)
        self.suffix_entry.grid(row=0, column=2, sticky="w",
                               padx=(0, GAP_LG), pady=PAD_CARD_Y)

        # 开始按钮
        self.start_btn = ttk.Button(
            card, text="▶ 开始处理", style="Accent.TButton",
            command=self.start_processing)
        self.start_btn.grid(row=0, column=3, sticky="e",
                            padx=(0, PAD_CARD), pady=PAD_CARD_Y)

        self._op_buttons = (self.add_files_btn, self.add_folder_btn,
                            self.clear_btn, self.start_btn)

    # ----------------------------------------------------------
    #  日志区域卡片
    # ----------------------------------------------------------
    def _build_log_section(self):
        card = tk.Frame(self.root, bg=COLOR_CARD_BG,
                        highlightbackground=COLOR_CARD_BORDER,
                        highlightthickness=1)
        card.grid(row=4, column=0, sticky="nsew",
                  padx=PAD_PAGE, pady=(0, PAD_PAGE))
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        title_row = tk.Frame(card, bg=COLOR_CARD_BG)
        title_row.grid(row=0, column=0, sticky="ew",
                       padx=PAD_CARD, pady=(PAD_TITLE_Y, GAP_MD))
        ttk.Label(title_row, text="处理日志",
                  style="Section.TLabel").pack(side=tk.LEFT)

        log_frame = tk.Frame(card, bg=COLOR_LOG_BORDER)
        log_frame.grid(row=1, column=0, sticky="nsew",
                       padx=PAD_CARD, pady=(0, PAD_CARD_Y))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            width=80, height=LOG_HEIGHT,
            state=tk.DISABLED,
            bg=COLOR_LOG_BG, fg=COLOR_TEXT,
            font=_make_font(FONT_SIZE_NORMAL, family=FONT_MONO),
            insertbackground=COLOR_TEXT,
            borderwidth=0, highlightthickness=0,
            relief="flat", wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

    # ============================================================
    #  业务逻辑
    # ============================================================
    def add_files(self):
        """添加选中的excel文件"""
        paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.last_dir)
        if paths:
            self.last_dir = os.path.dirname(paths[0])
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, p)
        self._update_file_count()

    def add_folder(self):
        """添加文件夹中的所有excel文件"""
        folder = filedialog.askdirectory(title="选择文件夹", initialdir=self.last_dir)
        if folder:
            self.last_dir = folder
        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".xlsx"):
                    full = os.path.join(root, f)
                    if full not in self.files:
                        self.files.append(full)
                        self.listbox.insert(tk.END, full)
        self._update_file_count()

    def _update_file_count(self):
        cnt = len(self.files)
        self._file_count_label.config(text=f"（{cnt} 个文件）")
        color = COLOR_PRIMARY if cnt > 0 else COLOR_TEXT_LIGHT
        self._file_count_label.config(foreground=color)

    def clear_list(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self._update_file_count()

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _set_buttons_state(self, state):
        for btn in self._op_buttons:
            btn.config(state=state)

    def start_processing(self):
        if not self.files:
            messagebox.showwarning("警告", "没有待处理的文件，请先添加文件或文件夹。")
            return
        self._set_buttons_state(tk.DISABLED)
        tid = self.template_var.get()
        t = tmpl.get_template_by_id(tid)
        self.log(f"使用模板：{t['name']}")
        self.log("开始处理...")
        thread = threading.Thread(target=self.process_all, daemon=True)
        thread.start()

    def process_all(self):
        tid = self.template_var.get()
        success_count = 0
        fail_count = 0
        for idx, filepath in enumerate(self.files, 1):
            self.log(f"[{idx}/{len(self.files)}] 正在处理: {os.path.basename(filepath)}")
            try:
                out_path = tmpl.process_file(filepath, self, tid)
                self.log(f"✓ 处理成功 -> {out_path}")
                success_count += 1
            except Exception as e:
                self.log(f"✗ 处理失败: {str(e)}")
                fail_count += 1
        self.log(f"处理完成！成功: {success_count}, 失败: {fail_count}")
        self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
