from datetime import datetime
from pathlib import Path
import re
import sys
import tkinter.font as tkfont

import customtkinter as ctk
from tkinter import filedialog, messagebox

from app_config import AppConfig, load_config, save_config
from export_utils import export_to_pdf, export_to_word
from text_cleaner import CleanOptions, clean_text, punctuation_consistency_warnings


def _enable_dpi_awareness() -> None:
    """Must be called BEFORE any Tk window is created."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


APP_NAME = "Reflow"
APP_TAGLINE = "AI 文本格式优化"
ACCENT = "#3B82F6"
ACCENT_HOVER = "#2563EB"
ACCENT_SECONDARY = "#6366F1"
CANVAS_BG = ("#F9FAFB", "#0F172A")
SIDEBAR_BG = ("#FFFFFF", "#1E293B")
CARD_BG = ("#FFFFFF", "#1E293B")
CARD_BORDER = ("#E5E7EB", "#334155")
INPUT_BG = ("#F9FAFB", "#172033")
INPUT_BORDER = ("#E5E7EB", "#334155")
TEXT_PRIMARY = ("#111827", "#F8FAFC")
TEXT_MUTED = ("#6B7280", "#94A3B8")
TEXT_SOFT = ("#9CA3AF", "#CBD5E1")
TABLE_GRID = ("#F3F4F6", "#334155")
TABLE_HEADER_BG = ("#F8FAFC", "#334155")
TABLE_ROW_BG = ("#FFFFFF", "#1E293B")
TABLE_ROW_ALT_BG = ("#F9FAFB", "#2A3B52")
SHADOW = ("0 2px 8px 0 rgba(0, 0, 0, 0.08)", "0 2px 8px 0 rgba(0, 0, 0, 0.4)")
UI_FONT_CANDIDATES = [
    "Segoe UI Variable",
    "Segoe UI",
    "Microsoft YaHei UI",
    "微软雅黑",
    "PingFang SC",
    "Noto Sans CJK SC",
    "Arial",
]


class CleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.config_model = load_config()
        self.history: list[dict[str, str]] = []
        self.history_index = -1
        self.history_limit = 30
        self.history_window = None

        self._configure_dpi()
        ctk.set_appearance_mode(self.config_model.theme_mode)
        ctk.set_default_color_theme("blue")

        self.title(f"{APP_NAME} · Pro")
        self._apply_window_geometry()
        self.configure(fg_color=CANVAS_BG)
        self.attributes('-alpha', 0.0)
        self.after(10, self._fade_in)

        self.ui_font_family = self._resolve_font(UI_FONT_CANDIDATES, "Arial")
        self._configure_tk_fonts()

        self.var_remove_markdown = ctk.BooleanVar(value=self.config_model.remove_markdown)
        self.var_normalize_punctuation = ctk.BooleanVar(value=self.config_model.normalize_punctuation)
        self.var_normalize_whitespace = ctk.BooleanVar(value=self.config_model.normalize_whitespace)
        self.var_merge_lines = ctk.BooleanVar(value=self.config_model.merge_lines)
        self.var_remove_emoji = ctk.BooleanVar(value=self.config_model.remove_emoji)
        self.var_indent_paragraph = ctk.BooleanVar(value=self.config_model.indent_paragraph)
        self.var_keep_tables = ctk.BooleanVar(value=self.config_model.keep_tables)

        self.preview_mode = "publish"

        self.status_text = ctk.StringVar(value="准备就绪")
        self.input_count_var = ctk.StringVar(value="输入 0 字")
        self.output_count_var = ctk.StringVar(value="输出 0 字")
        self.score_text = ctk.StringVar(value="质量评分：0")
        self.quality_badge = ctk.StringVar(value="未处理")
        self.punct_text = ctk.StringVar(value="标点一致性：未检测")

        self._build_ui()
        self._bind_events()
        self._record_state("启动")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _resolve_font(self, candidates: list[str], default: str) -> str:
        available = {name.lower(): name for name in tkfont.families(self)}
        for candidate in candidates:
            if candidate.lower() in available:
                return available[candidate.lower()]
        return default

    def _configure_dpi(self) -> None:
        try:
            self.update_idletasks()

            logical_dpi = 96
            if sys.platform.startswith("win"):
                try:
                    import ctypes
                    logical_dpi = ctypes.windll.user32.GetDpiForWindow(self.winfo_id())
                    if logical_dpi == 0:
                        logical_dpi = 96
                except Exception:
                    logical_dpi = 96

            self.tk.call('tk', 'scaling', logical_dpi / 72.0)
        except Exception:
            pass

    def _get_scaled_font_size(self, base_size: int) -> int:
        return base_size

    def _configure_tk_fonts(self) -> None:
        defaults = {
            "TkDefaultFont": self._get_scaled_font_size(12),
            "TkTextFont": self._get_scaled_font_size(14),
            "TkMenuFont": self._get_scaled_font_size(12),
            "TkHeadingFont": self._get_scaled_font_size(16),
            "TkCaptionFont": self._get_scaled_font_size(13),
        }
        for name, size in defaults.items():
            try:
                f = tkfont.nametofont(name)
                f.configure(family=self.ui_font_family, size=size)
            except Exception:
                pass

    def _parse_geometry(self, geometry: str, fallback: tuple[int, int]) -> tuple[int, int]:
        match = re.match(r"^(\d+)x(\d+)", geometry)
        if not match:
            return fallback
        return int(match.group(1)), int(match.group(2))

    def _fade_in(self) -> None:
        alpha = self.attributes('-alpha')
        if alpha < 1.0:
            alpha += 0.05
            self.attributes('-alpha', alpha)
            self.after(20, self._fade_in)

    def _apply_window_geometry(self) -> None:
        requested = (self.config_model.window_geometry or "").strip()
        req_width, req_height = self._parse_geometry(requested, (1280, 820))

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        max_width = max(1, int(screen_width * 0.92))
        max_height = max(1, int(screen_height * 0.9))
        min_width = min(1180, max_width)
        min_height = min(720, max_height)

        width = min(req_width, max_width)
        height = min(req_height, max_height)
        width = max(width, min_width)
        height = max(height, min_height)
        self.minsize(min_width, min_height)
        x = max(20, (screen_width - width) // 2)
        y = max(20, (screen_height - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _center_window(self) -> None:
        try:
            self.update_idletasks()
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            width = self.winfo_width()
            height = self.winfo_height()
            if width <= 1 or height <= 1:
                return
            x = max(20, (screen_width - width) // 2)
            y = max(20, (screen_height - height) // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            pass

    def _font(self, size: int, weight: str = "normal") -> ctk.CTkFont:
        return ctk.CTkFont(family=self.ui_font_family, size=size, weight=weight)

    def _text_font(self, size: int = 14) -> ctk.CTkFont:
        return ctk.CTkFont(family=self.ui_font_family, size=size)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        ghost_style = dict(
            fg_color=("#F3F4F6", "#334155"),
            hover_color=("#E5E7EB", "#475569"),
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=CARD_BORDER,
            font=self._font(12),
            corner_radius=8,
        )
        history_style = {**ghost_style, "font": self._font(11)}

        sidebar = ctk.CTkFrame(
            self,
            corner_radius=0,
            fg_color=SIDEBAR_BG,
            border_width=1,
            border_color=CARD_BORDER,
            width=340,
        )
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_rowconfigure(1, weight=1)

        sidebar_header = ctk.CTkFrame(sidebar, fg_color="transparent")
        sidebar_header.grid(row=0, column=0, sticky="ew", padx=22, pady=(16, 6))
        sidebar_header.grid_columnconfigure(0, weight=1)

        sidebar_scroll = ctk.CTkScrollableFrame(
            sidebar,
            fg_color="transparent",
            scrollbar_button_color=("#D8D2C6", "#2D3544"),
            scrollbar_button_hover_color=("#C9C1B3", "#394355"),
        )
        sidebar_scroll.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        sidebar_scroll.grid_columnconfigure(0, weight=1)

        brand = ctk.CTkFrame(sidebar_header, fg_color="transparent")
        brand.pack(fill="x")
        ctk.CTkLabel(brand, text=APP_NAME, font=self._font(32, "bold"), text_color=ACCENT).pack(anchor="w")
        ctk.CTkLabel(brand, text=APP_TAGLINE, font=self._font(13), text_color=TEXT_MUTED).pack(
            anchor="w", pady=(4, 0)
        )

        score_card = ctk.CTkFrame(
            sidebar_scroll, corner_radius=12, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER
        )
        score_card.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkLabel(score_card, textvariable=self.score_text, font=self._font(13, "bold"), text_color=TEXT_PRIMARY).pack(
            anchor="w", padx=12, pady=(10, 2)
        )
        ctk.CTkLabel(score_card, textvariable=self.quality_badge, font=self._font(12), text_color=TEXT_SOFT).pack(
            anchor="w", padx=12, pady=(0, 4)
        )
        ctk.CTkLabel(score_card, textvariable=self.punct_text, font=self._font(11), text_color=TEXT_MUTED).pack(
            anchor="w", padx=12, pady=(0, 10)
        )

        self.theme_switch = ctk.CTkSwitch(
            sidebar_scroll,
            text="深色模式",
            command=self.toggle_theme,
            font=self._font(11),
            text_color=TEXT_PRIMARY,
            progress_color=ACCENT,
        )
        self.theme_switch.pack(anchor="w", padx=12, pady=(4, 12))
        if self.config_model.theme_mode == "dark":
            self.theme_switch.select()

        self._section_title(sidebar_scroll, "清洗选项")
        self._add_switch(sidebar_scroll, "移除 Markdown", self.var_remove_markdown)
        self._add_switch(sidebar_scroll, "统一标点", self.var_normalize_punctuation)
        self._add_switch(sidebar_scroll, "空白规范化", self.var_normalize_whitespace)
        self._add_switch(sidebar_scroll, "合并断行", self.var_merge_lines)
        self._add_switch(sidebar_scroll, "段首空两格", self.var_indent_paragraph)
        self._add_switch(sidebar_scroll, "移除 Emoji", self.var_remove_emoji)
        self._add_switch(sidebar_scroll, "保留表格", self.var_keep_tables)

        ctk.CTkButton(
            sidebar_scroll,
            text="⚡ 一键优化  Ctrl+Enter",
            command=self.process_text,
            height=44,
            font=self._font(14, "bold"),
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="#ffffff",
            corner_radius=10,
            border_width=0,
        ).pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkButton(
            sidebar_scroll,
            text="📋 复制结果",
            command=self.copy_output,
            fg_color="transparent",
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=CARD_BORDER,
            hover_color=("#E5E7EB", "#475569"),
            height=36,
            font=self._font(12),
            corner_radius=8,
        ).pack(fill="x", padx=12, pady=(0, 12))

        self._section_title(sidebar_scroll, "历史")
        ctk.CTkButton(
            sidebar_scroll,
            text="↩️ 撤销",
            command=self.undo,
            height=30,
            **history_style,
        ).pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkButton(
            sidebar_scroll,
            text="↪️ 重做",
            command=self.redo,
            height=30,
            **history_style,
        ).pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkButton(
            sidebar_scroll,
            text="🕘 历史记录",
            command=self.show_history,
            height=32,
            **history_style,
        ).pack(fill="x", padx=12, pady=(0, 12))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=24, pady=24)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        toolbar = ctk.CTkFrame(main, corner_radius=14, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        toolbar.grid_columnconfigure(0, weight=1)
        toolbar.grid_columnconfigure(1, weight=0)

        left_group = ctk.CTkFrame(toolbar, fg_color="transparent")
        left_group.grid(row=0, column=0, sticky="w", padx=10, pady=8)
        right_group = ctk.CTkFrame(toolbar, fg_color="transparent")
        right_group.grid(row=0, column=1, sticky="e", padx=10, pady=8)

        ctk.CTkButton(left_group, text="📂 打开", command=self.open_file, width=82, **ghost_style).pack(
            side="left", padx=(0, 6)
        )
        ctk.CTkButton(left_group, text="💾 保存TXT", command=self.save_text, width=90, **ghost_style).pack(
            side="left", padx=6
        )
        ctk.CTkButton(left_group, text="📝 Word", command=self.export_word, width=82, **ghost_style).pack(
            side="left", padx=6
        )
        ctk.CTkButton(left_group, text="📕 PDF", command=self.export_pdf, width=82, **ghost_style).pack(
            side="left", padx=6
        )
        ctk.CTkButton(left_group, text="🧪 示例", command=self.fill_sample, width=82, **ghost_style).pack(
            side="left", padx=6
        )

        preview_switch = ctk.CTkSegmentedButton(
            right_group,
            values=["编辑视图", "发布预览"],
            command=self._switch_preview,
            corner_radius=10,
            font=self._font(12),
            selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=("#F3F4F6", "#334155"),
            unselected_hover_color=("#E5E7EB", "#475569"),
            text_color=TEXT_PRIMARY,
        )
        preview_switch.pack(side="right", padx=(6, 10), pady=10)
        preview_switch.set("编辑视图")

        self.editor_frame = ctk.CTkFrame(main, fg_color="transparent")
        self.editor_frame.grid(row=2, column=0, sticky="nsew")
        self.editor_frame.grid_columnconfigure(0, weight=1)
        self.editor_frame.grid_columnconfigure(1, weight=1)
        self.editor_frame.grid_rowconfigure(0, weight=1)

        input_card = ctk.CTkFrame(self.editor_frame, corner_radius=14, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
        input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        output_card = ctk.CTkFrame(self.editor_frame, corner_radius=14, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
        output_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        ctk.CTkLabel(input_card, text="原始文本", font=self._font(13, "bold"), text_color=TEXT_PRIMARY).pack(
            anchor="w", padx=12, pady=(10, 6)
        )
        self.input_text = ctk.CTkTextbox(
            input_card,
            corner_radius=10,
            font=self._text_font(15),
            fg_color=INPUT_BG,
            border_color=INPUT_BORDER,
            text_color=TEXT_PRIMARY,
        )
        self.input_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ctk.CTkLabel(output_card, text="优化结果", font=self._font(13, "bold"), text_color=TEXT_PRIMARY).pack(
            anchor="w", padx=12, pady=(10, 6)
        )
        self.output_text = ctk.CTkTextbox(
            output_card,
            corner_radius=10,
            font=self._text_font(15),
            fg_color=INPUT_BG,
            border_color=INPUT_BORDER,
            text_color=TEXT_PRIMARY,
        )
        self.output_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        try:
            self.input_text._textbox.configure(spacing1=4, spacing3=6)
            self.output_text._textbox.configure(spacing1=4, spacing3=6)
            self.input_text._textbox.configure(undo=True, maxundo=200, autoseparators=True)
            self.output_text._textbox.configure(undo=True, maxundo=200, autoseparators=True)
        except Exception:
            pass

        self.preview_card = ctk.CTkFrame(main, corner_radius=14, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
        ctk.CTkLabel(self.preview_card, text="发布预览", font=self._font(13, "bold"), text_color=TEXT_PRIMARY).pack(
            anchor="w", padx=12, pady=(10, 6)
        )
        self.preview_scroll = ctk.CTkScrollableFrame(
            self.preview_card,
            corner_radius=10,
            fg_color=INPUT_BG,
            border_width=1,
            border_color=INPUT_BORDER,
            scrollbar_button_color=("#D8D2C6", "#2D3544"),
            scrollbar_button_hover_color=("#C9C1B3", "#394355"),
        )
        self.preview_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        status_bar = ctk.CTkFrame(main, height=32, fg_color="transparent")
        status_bar.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        status_left = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_left.pack(side="left")

        self.status_dot = ctk.CTkLabel(
            status_left, text="\u25CF", font=self._font(8),
            text_color=("#22C55E", "#4ADE80"), width=14,
        )
        self.status_dot.pack(side="left", padx=(4, 4))
        ctk.CTkLabel(
            status_left, textvariable=self.status_text,
            font=self._font(11), text_color=TEXT_MUTED,
        ).pack(side="left")

        status_right = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_right.pack(side="right")

        ctk.CTkLabel(
            status_right, textvariable=self.input_count_var,
            font=self._font(11), text_color=TEXT_MUTED,
        ).pack(side="left")
        ctk.CTkLabel(
            status_right, text="|",
            font=self._font(11), text_color=TEXT_SOFT, width=20,
        ).pack(side="left", padx=(6, 6))
        ctk.CTkLabel(
            status_right, textvariable=self.output_count_var,
            font=self._font(11), text_color=TEXT_MUTED,
        ).pack(side="left", padx=(0, 4))

    def _section_title(self, parent: ctk.CTkFrame, text: str) -> None:
        ctk.CTkLabel(parent, text=text, font=self._font(11, "bold"), text_color=TEXT_SOFT).pack(
            anchor="w", padx=22, pady=(2, 8)
        )

    def _add_switch(self, parent: ctk.CTkFrame, text: str, variable: ctk.BooleanVar) -> None:
        ctk.CTkSwitch(
            parent,
            text=text,
            variable=variable,
            font=self._font(11),
            text_color=TEXT_PRIMARY,
            progress_color=ACCENT,
        ).pack(anchor="w", padx=22, pady=(0, 8))

    def _bind_events(self) -> None:
        self.bind("<Control-Return>", lambda _: self.process_text())
        self.bind("<Control-Shift-C>", lambda _: self.copy_output())
        self.bind("<Control-s>", lambda _: self.save_text())
        self.input_text.bind("<KeyRelease>", lambda _: self._update_stats())
        self.bind_all("<Control-z>", self._handle_undo, add="+")
        self.bind_all("<Control-y>", self._handle_redo, add="+")

    def _switch_preview(self, value: str) -> None:
        if value == "发布预览":
            self.preview_mode = "publish"
            self.editor_frame.grid_forget()
            self.preview_card.grid(row=2, column=0, sticky="nsew")
            self._render_publish_preview()
            self._set_status("已切换到发布预览")
            return

        self.preview_mode = "edit"
        self.preview_card.grid_forget()
        self.editor_frame.grid(row=2, column=0, sticky="nsew")
        self._set_status("已切换到编辑视图")

    def toggle_theme(self) -> None:
        mode = "dark" if self.theme_switch.get() == 1 else "light"
        ctk.set_appearance_mode(mode)
        self._set_status(f"主题已切换为{'深色' if mode == 'dark' else '浅色'}")

    def current_options(self) -> CleanOptions:
        return CleanOptions(
            remove_markdown=self.var_remove_markdown.get(),
            normalize_punctuation=self.var_normalize_punctuation.get(),
            normalize_whitespace=self.var_normalize_whitespace.get(),
            merge_wrapped_lines=self.var_merge_lines.get(),
            remove_emoji=self.var_remove_emoji.get(),
            indent_paragraph=self.var_indent_paragraph.get(),
            keep_tables=self.var_keep_tables.get(),
        )

    def _set_status(self, text: str) -> None:
        self.status_text.set(f"{datetime.now().strftime('%H:%M:%S')}  {text}")

    def _update_stats(self) -> None:
        in_len = len(self.input_text.get("1.0", "end-1c"))
        out_len = len(self.output_text.get("1.0", "end-1c"))
        self.input_count_var.set(f"输入 {in_len} 字")
        self.output_count_var.set(f"输出 {out_len} 字")

    def _quality_score(self, text: str) -> int:
        score = 100
        score -= min(20, len(re.findall(r"\s{3,}", text)) * 2)
        score -= min(20, len(re.findall(r"[\u200b\u200c\u200d\ufeff]", text)) * 3)
        score -= min(20, text.count("###") * 2 + text.count("```") * 3)
        if len(text.strip()) < 30:
            score -= 10
        return max(0, min(100, score))

    def _update_quality(self, text: str) -> None:
        score = self._quality_score(text)
        self.score_text.set(f"质量评分：{score}")
        if score >= 90:
            self.quality_badge.set("状态：可直接交付")
        elif score >= 75:
            self.quality_badge.set("状态：较好，建议复查")
        else:
            self.quality_badge.set("状态：需进一步优化")

    def _update_punctuation_consistency(self, text: str) -> None:
        warnings = punctuation_consistency_warnings(text)
        if warnings:
            self.punct_text.set(f"标点一致性：{', '.join(warnings)}")
        else:
            self.punct_text.set("标点一致性：良好")

    def _render_publish_preview(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        self._clear_preview()
        if not text.strip():
            self._add_preview_paragraph("（暂无内容，请先执行一键优化）")
            return

        for block in self._parse_preview_blocks(text):
            if block["type"] == "text":
                self._add_preview_paragraph(block["text"])
            elif block["type"] == "table":
                self._add_preview_table(block["rows"], block["has_header"])
            self._add_preview_spacing(8)

    def _clear_preview(self) -> None:
        for widget in self.preview_scroll.winfo_children():
            widget.destroy()

    def _add_preview_spacing(self, height: int) -> None:
        spacer = ctk.CTkFrame(self.preview_scroll, height=height, fg_color="transparent")
        spacer.pack(fill="x")

    def _add_preview_paragraph(self, text: str) -> None:
        label = ctk.CTkLabel(
            self.preview_scroll,
            text=text,
            justify="left",
            anchor="w",
            text_color=TEXT_PRIMARY,
            font=self._text_font(16),
            wraplength=self._preview_wraplength(),
        )
        label.pack(fill="x", padx=12, pady=(6, 6))

    def _preview_wraplength(self) -> int:
        width = self.preview_card.winfo_width()
        if width <= 1:
            return 920
        return max(360, width - 60)

    def _add_preview_table(self, rows: list[list[str]], has_header: bool) -> None:
        if not rows:
            return

        col_count = max(len(row) for row in rows)
        rows = [row + [""] * (col_count - len(row)) for row in rows]

        table_font = self._font(12)
        header_font = self._font(12, "bold")
        measure_font = tkfont.Font(family=self.ui_font_family, size=12, weight="normal")
        measure_header = tkfont.Font(family=self.ui_font_family, size=12, weight="bold")

        widths = [0] * col_count
        for row_index, row in enumerate(rows):
            for idx, cell in enumerate(row):
                measure = measure_header if has_header and row_index == 0 else measure_font
                widths[idx] = max(widths[idx], measure.measure(cell))

        widths = [min(max(w + 32, 96), 320) for w in widths]

        numeric_columns = []
        for col in range(col_count):
            numeric_hits = 0
            total_hits = 0
            for row_index, row in enumerate(rows):
                if has_header and row_index == 0:
                    continue
                value = row[col].strip()
                if not value:
                    continue
                total_hits += 1
                if re.fullmatch(r"[\d.,%+\-–/]+", value):
                    numeric_hits += 1
            numeric_columns.append(total_hits > 0 and numeric_hits / total_hits >= 0.6)

        table_frame = ctk.CTkFrame(
            self.preview_scroll,
            fg_color=TABLE_GRID,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=8,
        )
        table_frame.pack(fill="x", padx=6, pady=(2, 2))

        for col in range(col_count):
            table_frame.grid_columnconfigure(col, weight=1, minsize=widths[col])

        for row_index, row_cells in enumerate(rows):
            is_header = has_header and row_index == 0
            if is_header:
                row_bg = TABLE_HEADER_BG
            else:
                row_bg = TABLE_ROW_ALT_BG if row_index % 2 == 1 else TABLE_ROW_BG
            for col_index, cell in enumerate(row_cells):
                cell_frame = ctk.CTkFrame(
                    table_frame,
                    fg_color=row_bg,
                    corner_radius=0,
                )
                cell_frame.grid(row=row_index, column=col_index, sticky="nsew", padx=1, pady=1)
                cell_label = ctk.CTkLabel(
                    cell_frame,
                    text=cell,
                    text_color=TEXT_PRIMARY,
                    fg_color="transparent",
                    corner_radius=0,
                    anchor="e" if numeric_columns[col_index] else "w",
                    font=header_font if is_header else table_font,
                    wraplength=max(widths[col_index] - 30, 90),
                    justify="right" if numeric_columns[col_index] else "left",
                )
                cell_label.pack(fill="both", expand=True, padx=12, pady=8)

    def _parse_preview_blocks(self, text: str) -> list[dict[str, object]]:
        lines = text.split("\n")
        blocks: list[dict[str, object]] = []
        buffer: list[str] = []
        index = 0

        while index < len(lines):
            line = lines[index]
            if self._is_table_row_line(line):
                if buffer:
                    blocks.append({"type": "text", "text": "\n".join(buffer).strip()})
                    buffer = []

                table_lines = []
                while index < len(lines) and self._is_table_row_line(lines[index]):
                    table_lines.append(lines[index])
                    index += 1

                rows, has_header = self._parse_table_lines(table_lines)
                blocks.append({"type": "table", "rows": rows, "has_header": has_header})
                continue

            buffer.append(line)
            index += 1

        if buffer:
            blocks.append({"type": "text", "text": "\n".join(buffer).strip()})

        return [block for block in blocks if block.get("text") or block.get("rows")]

    def _parse_table_lines(self, lines: list[str]) -> tuple[list[list[str]], bool]:
        rows = []
        has_header = False
        for line in lines:
            if self._is_table_separator_line(line):
                has_header = True
                continue
            rows.append(self._split_table_row(line))
        return rows, has_header

    def _split_table_row(self, line: str) -> list[str]:
        stripped = line.strip().strip("|")
        parts = [part.strip() for part in stripped.split("|")]
        return parts

    def _is_table_row_line(self, line: str) -> bool:
        if "|" not in line:
            return False
        parts = [part.strip() for part in line.split("|") if part.strip()]
        return len(parts) >= 2

    def _is_table_separator_line(self, line: str) -> bool:
        stripped = line.strip()
        if "|" not in stripped:
            return False
        cells = [cell.strip() for cell in stripped.strip("|").split("|")]
        if not cells:
            return False
        for cell in cells:
            if not cell:
                return False
            if not re.fullmatch(r":?-{3,}:?", cell):
                return False
        return True

    def process_text(self) -> None:
        raw = self.input_text.get("1.0", "end")
        if not raw.strip():
            messagebox.showwarning("提示", "请先输入待处理文本")
            return

        result = clean_text(raw, self.current_options())
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", result)
        self._update_stats()
        self._update_quality(result)
        self._update_punctuation_consistency(result)
        if self.preview_mode == "publish":
            self._render_publish_preview()
        self._set_status("排版优化完成")
        self._record_state("一键优化")

    def copy_output(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("提示", "当前没有可复制内容")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self._set_status("结果已复制到剪贴板")

    def _dialog_dir(self, save: bool) -> str:
        value = self.config_model.last_save_dir if save else self.config_model.last_open_dir
        if value and Path(value).exists():
            return value
        return str(Path.cwd())

    def _default_filename(self, ext: str) -> str:
        return f"reflow_{datetime.now().strftime('%Y%m%d_%H%M')}.{ext}"

    def open_file(self) -> None:
        path = filedialog.askopenfilename(
            title="打开文本",
            initialdir=self._dialog_dir(False),
            filetypes=[("Text", "*.txt *.md *.text"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            content = Path(path).read_text(encoding="utf-8")
        except Exception as error:
            messagebox.showerror("读取失败", str(error))
            return
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", content)
        self.config_model.last_open_dir = str(Path(path).parent)
        self._update_stats()
        self._set_status(f"已打开：{Path(path).name}")
        self._record_state("打开文件")

    def save_text(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("提示", "请先生成优化结果")
            return
        path = filedialog.asksaveasfilename(
            title="保存文本",
            initialdir=self._dialog_dir(True),
            initialfile=self._default_filename("txt"),
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("Markdown", "*.md")],
        )
        if not path:
            return
        try:
            Path(path).write_text(text, encoding="utf-8")
        except Exception as error:
            messagebox.showerror("保存失败", str(error))
            return
        self.config_model.last_save_dir = str(Path(path).parent)
        self._set_status(f"已保存：{Path(path).name}")

    def export_word(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("提示", "请先生成优化结果")
            return
        path = filedialog.asksaveasfilename(
            title="导出 Word",
            initialdir=self._dialog_dir(True),
            initialfile=self._default_filename("docx"),
            defaultextension=".docx",
            filetypes=[("Word", "*.docx")],
        )
        if not path:
            return
        try:
            export_to_word(text, path)
        except RuntimeError as error:
            messagebox.showerror("依赖缺失", str(error))
            return
        except Exception as error:
            messagebox.showerror("导出失败", str(error))
            return
        self.config_model.last_save_dir = str(Path(path).parent)
        self._set_status(f"Word 导出成功：{Path(path).name}")

    def export_pdf(self) -> None:
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("提示", "请先生成优化结果")
            return
        path = filedialog.asksaveasfilename(
            title="导出 PDF",
            initialdir=self._dialog_dir(True),
            initialfile=self._default_filename("pdf"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        try:
            export_to_pdf(text, path)
        except RuntimeError as error:
            messagebox.showerror("依赖缺失", str(error))
            return
        except Exception as error:
            messagebox.showerror("导出失败", str(error))
            return
        self.config_model.last_save_dir = str(Path(path).parent)
        self._set_status(f"PDF 导出成功：{Path(path).name}")

    def fill_sample(self) -> None:
        demo = """### 本周运营复盘（草稿）

> 下面是我让 AI 整理的内容，但格式有点乱，帮我优化成可直接发给老板的版本。

---

1) 本周整体数据：曝光 128,900；点击 7,321；转化 412   ，环比提升 12.7%。。
2) 主要问题：
素材更新频率不稳定，导致周三-周四点击率下滑；
部分投放计划预算分配过于集中在晚间时段。

3) 下周计划:
- 统一更新素材节奏（每2天一次）
- 重点优化前3秒开头文案
- 增加A/B测试组（标题、封面、CTA）

| 指标 | 本周 | 上周 | 环比 |
|---|---:|---:|---:|
| 曝光 | 128900 | 113400 | +13.7% |
| 点击 | 7321 | 6880 | +6.4% |
| 转化 | 412 | 365 | +12.9% |

补充说明：详情见 [数据看板](https://example.com/dashboard) ，另有原始记录在附件里。

结论其实很简单  但原文太散了， 需要变成正式汇报格式。
"""
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", demo)
        self._update_stats()
        self._set_status("已载入示例文本")
        self._record_state("载入示例")

    def _capture_state(self, label: str) -> dict[str, str]:
        return {
            "label": label,
            "timestamp": datetime.now().strftime("%m-%d %H:%M"),
            "input": self.input_text.get("1.0", "end-1c"),
            "output": self.output_text.get("1.0", "end-1c"),
        }

    def _record_state(self, label: str) -> None:
        state = self._capture_state(label)
        if self.history and self.history_index >= 0:
            last = self.history[self.history_index]
            if last["input"] == state["input"] and last["output"] == state["output"]:
                return

        if self.history_index < len(self.history) - 1:
            self.history = self.history[: self.history_index + 1]

        self.history.append(state)
        if len(self.history) > self.history_limit:
            self.history.pop(0)
        self.history_index = len(self.history) - 1

    def _apply_state(self, state: dict[str, str]) -> None:
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", state["input"])
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", state["output"])
        self._update_stats()
        self._update_quality(state["output"])
        self._update_punctuation_consistency(state["output"])
        if self.preview_mode == "publish":
            self._render_publish_preview()

    def _handle_undo(self, event) -> str | None:
        if hasattr(event.widget, "edit_undo"):
            return None
        self.undo()
        return "break"

    def _handle_redo(self, event) -> str | None:
        if hasattr(event.widget, "edit_redo"):
            return None
        self.redo()
        return "break"

    def undo(self) -> None:
        if self.history_index <= 0:
            self._set_status("没有可撤销的记录")
            return
        self.history_index -= 1
        self._apply_state(self.history[self.history_index])
        self._set_status("已撤销到上一条记录")

    def redo(self) -> None:
        if self.history_index >= len(self.history) - 1:
            self._set_status("没有可重做的记录")
            return
        self.history_index += 1
        self._apply_state(self.history[self.history_index])
        self._set_status("已重做到下一条记录")

    def show_history(self) -> None:
        if self.history_window and self.history_window.winfo_exists():
            self.history_window.focus()
            return

        window = ctk.CTkToplevel(self)
        window.title("历史记录")
        window.geometry("560x520")
        window.configure(fg_color=CANVAS_BG)
        self.history_window = window

        header = ctk.CTkFrame(window, fg_color="transparent")
        header.pack(fill="x", padx=18, pady=(16, 6))
        ctk.CTkLabel(header, text="历史记录", font=self._font(16, "bold"), text_color=TEXT_PRIMARY).pack(
            side="left"
        )
        ctk.CTkLabel(
            header,
            text=f"{len(self.history)} 条",
            font=self._font(11),
            text_color=TEXT_MUTED,
        ).pack(side="left", padx=(8, 0))

        list_frame = ctk.CTkScrollableFrame(
            window,
            fg_color=CARD_BG,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=12,
        )
        list_frame.pack(fill="both", expand=True, padx=18, pady=(8, 18))

        for index in range(len(self.history) - 1, -1, -1):
            state = self.history[index]
            summary = self._history_summary(state)
            label = f"{state['timestamp']} · {state['label']}"
            button = ctk.CTkButton(
                list_frame,
                text=f"{label}\n{summary}",
                anchor="w",
                fg_color="transparent",
                hover_color=("#EFECE5", "#1E2530"),
                text_color=TEXT_PRIMARY,
                font=self._font(11),
                corner_radius=10,
                border_width=1,
                border_color=CARD_BORDER,
                command=lambda i=index: self._restore_history(i),
            )
            button.pack(fill="x", padx=8, pady=6)

    def _history_summary(self, state: dict[str, str]) -> str:
        text = state["output"] or state["input"]
        text = text.replace("\n", " ").strip()
        if not text:
            return "（空内容）"
        if len(text) > 80:
            text = text[:80] + "…"
        return text

    def _restore_history(self, index: int) -> None:
        if index < 0 or index >= len(self.history):
            return
        self.history_index = index
        self._apply_state(self.history[index])
        self._set_status("已恢复到选中历史记录")
        if self.history_window and self.history_window.winfo_exists():
            self.history_window.destroy()

    def _capture_config(self) -> AppConfig:
        theme_mode = "dark" if self.theme_switch.get() == 1 else "light"
        return AppConfig(
            remove_markdown=self.var_remove_markdown.get(),
            normalize_punctuation=self.var_normalize_punctuation.get(),
            normalize_whitespace=self.var_normalize_whitespace.get(),
            merge_lines=self.var_merge_lines.get(),
            remove_emoji=self.var_remove_emoji.get(),
            indent_paragraph=self.var_indent_paragraph.get(),
            keep_tables=self.var_keep_tables.get(),
            last_open_dir=self.config_model.last_open_dir,
            last_save_dir=self.config_model.last_save_dir,
            window_geometry=self.geometry(),
            theme_mode=theme_mode,
        )

    def on_close(self) -> None:
        try:
            save_config(self._capture_config())
        except Exception:
            pass
        self.destroy()


def run_gui() -> None:
    _enable_dpi_awareness()
    app = CleanerApp()
    app.mainloop()


if __name__ == "__main__":
    run_gui()
