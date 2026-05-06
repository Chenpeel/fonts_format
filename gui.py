#!/usr/bin/env python3
"""GUI 主界面（CustomTkinter）"""

import queue
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

import customtkinter as ctk

from core import DEFAULT_FONTS, HANDLERS, process_file


class FontFormatApp(ctk.CTk):

    ICONS = {'.pptx': '📊', '.docx': '📝', '.xlsx': '📈'}

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('System')
        ctk.set_default_color_theme('blue')

        self.title('Office 字体标准化工具')
        self.geometry('880x640')
        self.minsize(720, 540)

        self._files: list[Path] = []
        self._q: queue.Queue = queue.Queue()
        self._running = False

        self._build()
        self._poll()

    # ── 界面构建 ───────────────────────────────────────────────────────────────

    def _build(self):
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self._build_header()
        self._build_file_panel()
        self._build_settings_panel()
        self._build_progress_bar()
        self._build_log_panel()

    def _build_header(self):
        bar = ctk.CTkFrame(self, height=48, corner_radius=0)
        bar.grid(row=0, column=0, columnspan=2, sticky='ew')
        bar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            bar, text='Office 字体标准化工具',
            font=ctk.CTkFont(size=16, weight='bold'),
        ).grid(row=0, column=0, padx=16, pady=10, sticky='w')

        self._theme_btn = ctk.CTkButton(
            bar, text='🌙', width=36, height=28,
            command=self._toggle_theme,
        )
        self._theme_btn.grid(row=0, column=1, padx=12, sticky='e')

    def _build_file_panel(self):
        panel = ctk.CTkFrame(self)
        panel.grid(row=1, column=0, padx=(12, 6), pady=6, sticky='nsew')
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            panel, text='📂  待处理文件',
            font=ctk.CTkFont(weight='bold'),
        ).grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 4), sticky='w')

        self._file_scroll = ctk.CTkScrollableFrame(panel)
        self._file_scroll.grid(row=1, column=0, columnspan=2,
                                padx=8, pady=4, sticky='nsew')
        self._file_scroll.grid_columnconfigure(0, weight=1)

        btn_row = ctk.CTkFrame(panel, fg_color='transparent')
        btn_row.grid(row=2, column=0, columnspan=2,
                     padx=8, pady=(4, 10), sticky='ew')

        ctk.CTkButton(btn_row, text='＋ 添加文件', width=110,
                       command=self._add_files).pack(side='left', padx=4)
        ctk.CTkButton(btn_row, text='清空', width=60,
                       fg_color='gray40', hover_color='gray30',
                       command=self._clear_files).pack(side='left', padx=4)
        self._count_lbl = ctk.CTkLabel(btn_row, text='0 个文件',
                                        text_color='gray')
        self._count_lbl.pack(side='right', padx=8)

    def _build_settings_panel(self):
        panel = ctk.CTkFrame(self)
        panel.grid(row=1, column=1, padx=(6, 12), pady=6, sticky='nsew')
        panel.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            panel, text='⚙  字体规则',
            font=ctk.CTkFont(weight='bold'),
        ).grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 8), sticky='w')

        self._font_entries: dict[str, ctk.CTkEntry] = {}
        rows = [
            ('chinese',  '中文字体'),
            ('latin',    '英文字体'),
            ('japanese', '日文字体'),
        ]
        for i, (key, label) in enumerate(rows):
            ctk.CTkLabel(panel, text=label).grid(
                row=i + 1, column=0, padx=(12, 4), pady=5, sticky='w')
            e = ctk.CTkEntry(panel)
            e.insert(0, DEFAULT_FONTS[key])
            e.grid(row=i + 1, column=1, padx=(4, 12), pady=5, sticky='ew')
            self._font_entries[key] = e

        # 分隔
        ctk.CTkFrame(panel, height=1, fg_color='gray40').grid(
            row=4, column=0, columnspan=2, padx=12, pady=10, sticky='ew')

        ctk.CTkLabel(
            panel, text='📁  输出目录',
            font=ctk.CTkFont(weight='bold'),
        ).grid(row=5, column=0, columnspan=2, padx=12, pady=(0, 6), sticky='w')

        self._out_var = tk.IntVar(value=0)
        ctk.CTkRadioButton(
            panel, text='与原文件相同目录',
            variable=self._out_var, value=0,
            command=self._on_outmode,
        ).grid(row=6, column=0, columnspan=2, padx=16, pady=3, sticky='w')
        ctk.CTkRadioButton(
            panel, text='自定义目录',
            variable=self._out_var, value=1,
            command=self._on_outmode,
        ).grid(row=7, column=0, columnspan=2, padx=16, pady=3, sticky='w')

        custom_row = ctk.CTkFrame(panel, fg_color='transparent')
        custom_row.grid(row=8, column=0, columnspan=2,
                        padx=12, pady=(2, 12), sticky='ew')
        custom_row.grid_columnconfigure(0, weight=1)

        self._out_entry = ctk.CTkEntry(
            custom_row, placeholder_text='选择输出目录...', state='disabled')
        self._out_entry.grid(row=0, column=0, sticky='ew', padx=(0, 6))

        self._out_btn = ctk.CTkButton(
            custom_row, text='浏览', width=50, state='disabled',
            command=self._browse_out)
        self._out_btn.grid(row=0, column=1)

    def _build_progress_bar(self):
        bar = ctk.CTkFrame(self, height=52)
        bar.grid(row=2, column=0, columnspan=2, padx=12, pady=4, sticky='ew')
        bar.grid_columnconfigure(0, weight=1)

        self._progress = ctk.CTkProgressBar(bar)
        self._progress.set(0)
        self._progress.grid(row=0, column=0, padx=(12, 8), pady=14, sticky='ew')

        self._prog_lbl = ctk.CTkLabel(bar, text='准备就绪', width=90,
                                       text_color='gray')
        self._prog_lbl.grid(row=0, column=1, padx=4)

        self._start_btn = ctk.CTkButton(
            bar, text='▶  开始处理', width=130,
            command=self._start)
        self._start_btn.grid(row=0, column=2, padx=(4, 12))

    def _build_log_panel(self):
        panel = ctk.CTkFrame(self)
        panel.grid(row=3, column=0, columnspan=2,
                   padx=12, pady=(0, 12), sticky='nsew')
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(panel, fg_color='transparent')
        hdr.grid(row=0, column=0, sticky='ew', padx=10, pady=(6, 0))
        ctk.CTkLabel(hdr, text='📋  处理日志',
                     font=ctk.CTkFont(weight='bold')).pack(side='left')
        ctk.CTkButton(hdr, text='清空', width=50, height=24,
                       fg_color='gray40', hover_color='gray30',
                       command=self._clear_log).pack(side='right')

        self._log_box = ctk.CTkTextbox(
            panel, state='disabled',
            font=ctk.CTkFont(family='Courier', size=12))
        self._log_box.grid(row=1, column=0, padx=8, pady=(4, 8), sticky='nsew')

    # ── 事件 ───────────────────────────────────────────────────────────────────

    def _toggle_theme(self):
        mode = ctk.get_appearance_mode()
        ctk.set_appearance_mode('Light' if mode == 'Dark' else 'Dark')
        self._theme_btn.configure(text='☀️' if mode == 'Dark' else '🌙')

    def _on_outmode(self):
        is_custom = self._out_var.get() == 1
        state = 'normal' if is_custom else 'disabled'
        self._out_entry.configure(state=state)
        self._out_btn.configure(state=state)

    def _browse_out(self):
        d = filedialog.askdirectory(title='选择输出目录')
        if d:
            self._out_entry.configure(state='normal')
            self._out_entry.delete(0, 'end')
            self._out_entry.insert(0, d)
            if self._out_var.get() != 1:
                self._out_var.set(1)
                self._on_outmode()

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title='选择 Office 文件',
            filetypes=[
                ('Office 文件', '*.pptx *.docx *.xlsx'),
                ('PowerPoint',  '*.pptx'),
                ('Word',        '*.docx'),
                ('Excel',       '*.xlsx'),
                ('所有文件',    '*.*'),
            ],
        )
        for s in paths:
            p = Path(s)
            if p.suffix.lower() not in HANDLERS:
                continue
            if p not in self._files:
                self._files.append(p)
                self._add_file_row(p)
        self._refresh_count()

    def _add_file_row(self, path: Path):
        row = ctk.CTkFrame(self._file_scroll)
        row.pack(fill='x', padx=4, pady=2)
        row.grid_columnconfigure(1, weight=1)

        icon = self.ICONS.get(path.suffix.lower(), '📄')
        ctk.CTkLabel(row, text=icon, width=26).grid(
            row=0, column=0, padx=(6, 2))
        ctk.CTkLabel(row, text=path.name, anchor='w').grid(
            row=0, column=1, padx=4, pady=5, sticky='ew')
        ctk.CTkLabel(row, text=f'…/{path.parent.name}',
                     text_color='gray',
                     font=ctk.CTkFont(size=11)).grid(
            row=0, column=2, padx=6)
        ctk.CTkButton(
            row, text='✕', width=26, height=26,
            fg_color='transparent', hover_color='#c0392b',
            command=lambda r=row, p=path: self._remove_file(r, p),
        ).grid(row=0, column=3, padx=(2, 6))

    def _remove_file(self, row, path: Path):
        if path in self._files:
            self._files.remove(path)
        row.destroy()
        self._refresh_count()

    def _clear_files(self):
        self._files.clear()
        for w in self._file_scroll.winfo_children():
            w.destroy()
        self._refresh_count()

    def _refresh_count(self):
        n = len(self._files)
        self._count_lbl.configure(text=f'{n} 个文件')

    def _clear_log(self):
        self._log_box.configure(state='normal')
        self._log_box.delete('1.0', 'end')
        self._log_box.configure(state='disabled')

    # ── 处理 ───────────────────────────────────────────────────────────────────

    def _get_fonts(self) -> dict:
        return {k: e.get().strip() or DEFAULT_FONTS[k]
                for k, e in self._font_entries.items()}

    def _get_out_dir(self, file_path: Path) -> Path:
        if self._out_var.get() == 1:
            custom = self._out_entry.get().strip()
            if custom:
                return Path(custom)
        return file_path.parent

    def _start(self):
        if self._running:
            return
        if not self._files:
            messagebox.showwarning('提示', '请先添加要处理的文件')
            return
        if self._out_var.get() == 1 and not self._out_entry.get().strip():
            messagebox.showwarning('提示', '请选择自定义输出目录')
            return

        self._running = True
        self._start_btn.configure(state='disabled', text='处理中…')
        self._progress.set(0)
        self._prog_lbl.configure(text='0%')

        # 在主线程提前读取配置，避免子线程访问 Tk 控件
        fonts    = self._get_fonts()
        files    = list(self._files)
        out_dirs = {p: self._get_out_dir(p) for p in files}

        threading.Thread(
            target=self._worker, args=(files, out_dirs, fonts), daemon=True
        ).start()

    def _worker(self, files: list[Path], out_dirs: dict, fonts: dict):
        total = len(files)
        for i, p in enumerate(files):
            self._q.put(f'处理: {p.name}')
            try:
                out_dir = out_dirs[p]
                out_dir.mkdir(parents=True, exist_ok=True)
                out = process_file(p, out_dir, fonts,
                                   lambda msg: self._q.put(f'  {msg}'))
                self._q.put(f'✓  {p.name}  →  {out.name}')
            except Exception as e:
                self._q.put(f'✗  {p.name}: {e}')

            self._q.put(('__PROG__', (i + 1) / total,
                         f'{i + 1} / {total}'))

        self._q.put('__DONE__')

    def _poll(self):
        while not self._q.empty():
            msg = self._q.get_nowait()

            if msg == '__DONE__':
                self._running = False
                self._start_btn.configure(state='normal', text='▶  开始处理')
                self._prog_lbl.configure(text='完成 ✓')
                self._append_log('─' * 36)
                self._append_log('全部处理完成')

            elif isinstance(msg, tuple) and msg[0] == '__PROG__':
                _, val, label = msg
                self._progress.set(val)
                self._prog_lbl.configure(text=label)

            else:
                self._append_log(str(msg))

        self.after(80, self._poll)

    def _append_log(self, text: str):
        self._log_box.configure(state='normal')
        self._log_box.insert('end', text + '\n')
        self._log_box.see('end')
        self._log_box.configure(state='disabled')
