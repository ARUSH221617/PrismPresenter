"""
PPTX Jahat - Modern Black & Red Design System and UI Component Library.
Provides styling tokens, custom dark/red widgets, badges, cards, navigation, and terminal-style logs.
"""

from typing import Dict, Any, Optional, Callable, List
import tkinter as tk
from tkinter import ttk, font as tkfont


# ==========================================
# COLOR TOKENS: MODERN BLACK & CRIMSON RED
# ==========================================
class Theme:
    # Deep obsidian / charcoal backgrounds
    BG_DARKEST = "#0b0c10"     # Ultra deep canvas / window root
    BG_MAIN = "#12141a"        # Primary card and container surface
    BG_SURFACE = "#1a1d24"     # Elevated card / panel / input surface
    BG_SURFACE_HOVER = "#242833" # Hovered surface state
    BG_INPUT = "#15181f"       # Input field background
    BG_HEADER = "#0d0f14"      # App top bar / header banner

    # Crimson / Cyberpunk Red Accents
    RED_PRIMARY = "#e50914"    # Vibrant signature red (buttons, active states)
    RED_HOVER = "#ff2a38"      # Glow / Hover red
    RED_ACTIVE = "#b80710"     # Pressed red
    RED_MUTED = "#52161b"      # Muted red border / low-light badge
    RED_ACCENT_BG = "#2a1215"  # Subtle red tinted card/panel background
    RED_GLOW = "#ff4d5a"       # Bright highlight red

    # Neutral Borders and Dividers
    BORDER_DARK = "#262b36"    # Standard borders
    BORDER_LIGHT = "#363d4d"   # Active borders / subtle focus
    BORDER_RED = "#e50914"     # Active focus ring / accent border

    # High Contrast & Polished Typography Colors
    TEXT_WHITE = "#f8f9fa"     # Primary text / headings
    TEXT_MAIN = "#e2e8f0"      # Standard body text
    TEXT_MUTED = "#94a3b8"     # Secondary captions / hints
    TEXT_DIM = "#64748b"       # Disabled / very subtle labels
    TEXT_RED = "#ff4d5a"       # Red accent text / alerts / highlights
    TEXT_SUCCESS = "#4ade80"   # Success indicators (green accent)
    TEXT_WARNING = "#facc15"   # Warning indicators (amber accent)

    # Scrollbars & Sliders
    SCROLLBAR_TROUGH = "#12141a"
    SCROLLBAR_THUMB = "#2d3340"
    SCROLLBAR_ACTIVE = "#e50914"

    # Status / Badge Colors
    BADGE_BG_RED = "#3c1417"
    BADGE_BORDER_RED = "#ff2a38"
    BADGE_TEXT_RED = "#ff6b76"

    # Fonts
    FONT_FAMILY = "Segoe UI"
    FONT_CODE = "Consolas"
    FONT_HEADER = ("Segoe UI", 16, "bold")
    FONT_SUBHEADER = ("Segoe UI", 12, "bold")
    FONT_TITLE = ("Segoe UI", 10, "bold")
    FONT_BODY = ("Segoe UI", 9)
    FONT_BODY_BOLD = ("Segoe UI", 9, "bold")
    FONT_CAPTION = ("Segoe UI", 8)
    FONT_BUTTON = ("Segoe UI", 9, "bold")
    FONT_LOG = ("Consolas", 9)


def apply_theme_to_ttk(root: tk.Tk):
    """Configures the global TTK styles for the Black & Crimson Red Theme."""
    style = ttk.Style(root)
    style.theme_use("clam")

    # Base background
    root.configure(bg=Theme.BG_DARKEST)

    # TFrame
    style.configure(
        "TFrame",
        background=Theme.BG_MAIN
    )
    style.configure(
        "Root.TFrame",
        background=Theme.BG_DARKEST
    )
    style.configure(
        "Surface.TFrame",
        background=Theme.BG_SURFACE
    )
    style.configure(
        "Accent.TFrame",
        background=Theme.RED_ACCENT_BG
    )

    # TLabel
    style.configure(
        "TLabel",
        background=Theme.BG_MAIN,
        foreground=Theme.TEXT_MAIN,
        font=Theme.FONT_BODY
    )
    style.configure(
        "Surface.TLabel",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_MAIN,
        font=Theme.FONT_BODY
    )
    style.configure(
        "Header.TLabel",
        background=Theme.BG_DARKEST,
        foreground=Theme.TEXT_WHITE,
        font=Theme.FONT_HEADER
    )
    style.configure(
        "Subheader.TLabel",
        background=Theme.BG_DARKEST,
        foreground=Theme.TEXT_RED,
        font=Theme.FONT_SUBHEADER
    )
    style.configure(
        "Muted.TLabel",
        background=Theme.BG_MAIN,
        foreground=Theme.TEXT_MUTED,
        font=Theme.FONT_CAPTION
    )
    style.configure(
        "SurfaceMuted.TLabel",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_MUTED,
        font=Theme.FONT_CAPTION
    )
    style.configure(
        "Accent.TLabel",
        background=Theme.BG_MAIN,
        foreground=Theme.TEXT_RED,
        font=Theme.FONT_BODY_BOLD
    )

    # TLabelframe
    style.configure(
        "TLabelframe",
        background=Theme.BG_MAIN,
        foreground=Theme.TEXT_WHITE,
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.BORDER_DARK,
        darkcolor=Theme.BORDER_DARK,
        borderwidth=1,
        relief="solid",
        padding=10
    )
    style.configure(
        "TLabelframe.Label",
        background=Theme.BG_MAIN,
        foreground=Theme.TEXT_RED,
        font=Theme.FONT_TITLE
    )
    style.configure(
        "Surface.TLabelframe",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_WHITE,
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.BORDER_DARK,
        darkcolor=Theme.BORDER_DARK,
        borderwidth=1,
        relief="solid",
        padding=10
    )
    style.configure(
        "Surface.TLabelframe.Label",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_RED,
        font=Theme.FONT_TITLE
    )

    # TButton (Primary Crimson Red)
    style.configure(
        "TButton",
        background=Theme.RED_PRIMARY,
        foreground=Theme.TEXT_WHITE,
        font=Theme.FONT_BUTTON,
        bordercolor=Theme.RED_PRIMARY,
        lightcolor=Theme.RED_PRIMARY,
        darkcolor=Theme.RED_PRIMARY,
        focuscolor="none",
        relief="flat",
        padding=(12, 6)
    )
    style.map(
        "TButton",
        background=[
            ("pressed", Theme.RED_ACTIVE),
            ("active", Theme.RED_HOVER),
            ("disabled", Theme.BG_SURFACE)
        ],
        foreground=[
            ("disabled", Theme.TEXT_DIM),
            ("!disabled", Theme.TEXT_WHITE)
        ],
        bordercolor=[
            ("disabled", Theme.BORDER_DARK),
            ("!disabled", Theme.RED_HOVER)
        ]
    )

    # Secondary Outline Button
    style.configure(
        "Secondary.TButton",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_MAIN,
        font=Theme.FONT_BODY_BOLD,
        bordercolor=Theme.BORDER_LIGHT,
        lightcolor=Theme.BORDER_LIGHT,
        darkcolor=Theme.BORDER_LIGHT,
        focuscolor="none",
        relief="solid",
        padding=(10, 5)
    )
    style.map(
        "Secondary.TButton",
        background=[
            ("pressed", Theme.BG_DARKEST),
            ("active", Theme.BG_SURFACE_HOVER),
            ("disabled", Theme.BG_MAIN)
        ],
        foreground=[
            ("disabled", Theme.TEXT_DIM),
            ("active", Theme.TEXT_WHITE),
            ("!disabled", Theme.TEXT_MAIN)
        ],
        bordercolor=[
            ("active", Theme.RED_PRIMARY),
            ("disabled", Theme.BORDER_DARK)
        ]
    )

    # TEntry
    style.configure(
        "TEntry",
        fieldbackground=Theme.BG_INPUT,
        foreground=Theme.TEXT_WHITE,
        insertcolor=Theme.TEXT_RED,
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.BORDER_DARK,
        darkcolor=Theme.BORDER_DARK,
        relief="solid",
        borderwidth=1,
        padding=6
    )
    style.map(
        "TEntry",
        bordercolor=[
            ("focus", Theme.RED_PRIMARY),
            ("!focus", Theme.BORDER_DARK)
        ],
        fieldbackground=[
            ("disabled", Theme.BG_SURFACE),
            ("!disabled", Theme.BG_INPUT)
        ]
    )

    # TCombobox
    style.configure(
        "TCombobox",
        fieldbackground=Theme.BG_INPUT,
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_WHITE,
        arrowcolor=Theme.TEXT_RED,
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.BORDER_DARK,
        darkcolor=Theme.BORDER_DARK,
        relief="solid",
        borderwidth=1,
        padding=5
    )
    style.map(
        "TCombobox",
        fieldbackground=[
            ("readonly", Theme.BG_INPUT),
            ("disabled", Theme.BG_SURFACE)
        ],
        foreground=[
            ("readonly", Theme.TEXT_WHITE),
            ("disabled", Theme.TEXT_DIM)
        ],
        bordercolor=[
            ("focus", Theme.RED_PRIMARY),
            ("active", Theme.RED_HOVER),
            ("!focus", Theme.BORDER_DARK)
        ],
        arrowcolor=[
            ("active", Theme.RED_HOVER),
            ("!active", Theme.TEXT_RED)
        ]
    )

    # TNotebook (Tabs in Black & Red)
    style.configure(
        "TNotebook",
        background=Theme.BG_DARKEST,
        borderwidth=0,
        tabmargins=[4, 4, 4, 0]
    )
    style.configure(
        "TNotebook.Tab",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_MUTED,
        font=Theme.FONT_TITLE,
        padding=(16, 8),
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.BORDER_DARK,
        darkcolor=Theme.BORDER_DARK,
        relief="flat"
    )
    style.map(
        "TNotebook.Tab",
        background=[
            ("selected", Theme.RED_PRIMARY),
            ("active", Theme.BG_SURFACE_HOVER),
            ("!selected", Theme.BG_MAIN)
        ],
        foreground=[
            ("selected", Theme.TEXT_WHITE),
            ("active", Theme.TEXT_WHITE),
            ("!selected", Theme.TEXT_MUTED)
        ],
        bordercolor=[
            ("selected", Theme.RED_HOVER),
            ("!selected", Theme.BORDER_DARK)
        ]
    )

    # TProgressbar
    style.configure(
        "Horizontal.TProgressbar",
        background=Theme.RED_PRIMARY,
        troughcolor=Theme.BG_INPUT,
        bordercolor=Theme.BORDER_DARK,
        lightcolor=Theme.RED_PRIMARY,
        darkcolor=Theme.RED_PRIMARY,
        relief="flat"
    )

    # TPanedwindow
    style.configure(
        "TPanedwindow",
        background=Theme.BG_DARKEST
    )
    style.configure(
        "Sash",
        sashthickness=5,
        gripcount=0,
        background=Theme.BORDER_DARK
    )

    # TScrollbar
    style.configure(
        "Vertical.TScrollbar",
        troughcolor=Theme.SCROLLBAR_TROUGH,
        background=Theme.SCROLLBAR_THUMB,
        bordercolor=Theme.BORDER_DARK,
        arrowcolor=Theme.TEXT_MUTED,
        relief="flat"
    )
    style.map(
        "Vertical.TScrollbar",
        background=[
            ("active", Theme.RED_PRIMARY),
            ("pressed", Theme.RED_ACTIVE),
            ("!disabled", Theme.SCROLLBAR_THUMB)
        ]
    )
    style.configure(
        "Horizontal.TScrollbar",
        troughcolor=Theme.SCROLLBAR_TROUGH,
        background=Theme.SCROLLBAR_THUMB,
        bordercolor=Theme.BORDER_DARK,
        arrowcolor=Theme.TEXT_MUTED,
        relief="flat"
    )
    style.map(
        "Horizontal.TScrollbar",
        background=[
            ("active", Theme.RED_PRIMARY),
            ("pressed", Theme.RED_ACTIVE),
            ("!disabled", Theme.SCROLLBAR_THUMB)
        ]
    )

    # Treeview (for Modern File Lists & Catalogs)
    style.configure(
        "Treeview",
        background=Theme.BG_DARKEST,
        fieldbackground=Theme.BG_DARKEST,
        foreground=Theme.TEXT_MAIN,
        font=Theme.FONT_BODY,
        rowheight=28,
        borderwidth=0
    )
    style.map(
        "Treeview",
        background=[
            ("selected", Theme.RED_MUTED)
        ],
        foreground=[
            ("selected", Theme.TEXT_WHITE)
        ]
    )
    style.configure(
        "Treeview.Heading",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_RED,
        font=Theme.FONT_TITLE,
        relief="flat",
        padding=(8, 6)
    )
    style.map(
        "Treeview.Heading",
        background=[
            ("active", Theme.BG_SURFACE_HOVER)
        ],
        foreground=[
            ("active", Theme.TEXT_WHITE)
        ]
    )


# ==========================================
# MODERN CUSTOM CANVAS & UI WIDGETS
# ==========================================

class ModernCard(tk.Frame):
    """
    Sleek dark card container with subtle borders, optional red accent stripe,
    and structured header/body areas.
    """
    def __init__(
        self,
        parent,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        accent_color: Optional[str] = None,
        show_accent_stripe: bool = False,
        **kwargs
    ):
        bg = kwargs.pop("bg", Theme.BG_SURFACE)
        bd = kwargs.pop("bd", 1)
        relief = kwargs.pop("relief", "solid")
        highlightbackground = kwargs.pop("highlightbackground", Theme.BORDER_DARK)
        highlightthickness = kwargs.pop("highlightthickness", 1)

        super().__init__(
            parent,
            bg=bg,
            bd=bd,
            relief=relief,
            highlightbackground=highlightbackground,
            highlightthickness=highlightthickness,
            **kwargs
        )
        self.accent_color = accent_color or Theme.RED_PRIMARY
        
        if show_accent_stripe:
            stripe = tk.Frame(self, bg=self.accent_color, height=3)
            stripe.pack(fill=tk.X, side=tk.TOP)

        if title:
            header_frame = tk.Frame(self, bg=bg, padx=12, pady=8)
            header_frame.pack(fill=tk.X, side=tk.TOP)

            lbl_title = tk.Label(
                header_frame,
                text=title,
                bg=bg,
                fg=Theme.TEXT_WHITE,
                font=Theme.FONT_TITLE,
                anchor="w"
            )
            lbl_title.pack(side=tk.LEFT)

            if subtitle:
                lbl_sub = tk.Label(
                    header_frame,
                    text=subtitle,
                    bg=bg,
                    fg=Theme.TEXT_MUTED,
                    font=Theme.FONT_CAPTION,
                    anchor="w"
                )
                lbl_sub.pack(side=tk.LEFT, padx=(8, 0))

            sep = tk.Frame(self, bg=Theme.BORDER_DARK, height=1)
            sep.pack(fill=tk.X, side=tk.TOP)

        self.content_frame = tk.Frame(self, bg=bg, padx=12, pady=10)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

    @property
    def body(self) -> tk.Frame:
        return self.content_frame


class Badge(tk.Frame):
    """Modern pill badge with colored outline and text for status or tags."""
    def __init__(
        self,
        parent,
        text: str,
        bg_color: str = Theme.BADGE_BG_RED,
        fg_color: str = Theme.BADGE_TEXT_RED,
        border_color: str = Theme.BADGE_BORDER_RED,
        **kwargs
    ):
        super().__init__(
            parent,
            bg=bg_color,
            highlightbackground=border_color,
            highlightthickness=1,
            padx=8,
            pady=2,
            **kwargs
        )
        self.label = tk.Label(
            self,
            text=text,
            bg=bg_color,
            fg=fg_color,
            font=("Segoe UI", 8, "bold")
        )
        self.label.pack()

    def set_text(self, text: str, fg_color: Optional[str] = None, bg_color: Optional[str] = None):
        self.label.config(text=text)
        if fg_color:
            self.label.config(fg=fg_color)
        if bg_color:
            self.config(bg=bg_color)
            self.label.config(bg=bg_color)


class ConsoleLogWidget(tk.Frame):
    """
    Modern high-tech terminal console widget with styled colored lines,
    timestamp prefixes, auto-scroll, clear action, and copy support.
    """
    def __init__(self, parent, title: str = "Live Logs", height: int = 10, **kwargs):
        super().__init__(
            parent,
            bg=Theme.BG_DARKEST,
            highlightbackground=Theme.BORDER_DARK,
            highlightthickness=1,
            **kwargs
        )

        # Header bar
        header = tk.Frame(self, bg=Theme.BG_HEADER, padx=8, pady=4)
        header.pack(fill=tk.X, side=tk.TOP)

        # Terminal dot indicators (Red, Dark, Muted)
        dots_frame = tk.Frame(header, bg=Theme.BG_HEADER)
        dots_frame.pack(side=tk.LEFT, padx=(2, 6))

        for color in [Theme.RED_PRIMARY, "#f59e0b", "#10b981"]:
            dot = tk.Frame(dots_frame, bg=color, width=8, height=8)
            dot.pack(side=tk.LEFT, padx=2)

        lbl = tk.Label(
            header,
            text=title,
            bg=Theme.BG_HEADER,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_TITLE
        )
        lbl.pack(side=tk.LEFT, padx=4)

        # Action buttons on header right
        btn_clear = tk.Button(
            header,
            text="Clear",
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_MUTED,
            activebackground=Theme.BG_SURFACE_HOVER,
            activeforeground=Theme.TEXT_WHITE,
            relief="flat",
            font=("Segoe UI", 7, "bold"),
            padx=6,
            pady=1,
            command=self.clear
        )
        btn_clear.pack(side=tk.RIGHT, padx=2)

        sep = tk.Frame(self, bg=Theme.BORDER_DARK, height=1)
        sep.pack(fill=tk.X, side=tk.TOP)

        # Text Area with Custom Colors
        self.text_area = tk.Text(
            self,
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MAIN,
            insertbackground=Theme.RED_PRIMARY,
            selectbackground=Theme.RED_MUTED,
            selectforeground=Theme.TEXT_WHITE,
            font=Theme.FONT_LOG,
            wrap=tk.WORD,
            height=height,
            bd=0,
            padx=8,
            pady=8,
            state="disabled"
        )
        
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Setup Tag Styles
        self.text_area.tag_config("info", foreground=Theme.TEXT_MAIN)
        self.text_area.tag_config("success", foreground=Theme.TEXT_SUCCESS, font=("Consolas", 9, "bold"))
        self.text_area.tag_config("warning", foreground=Theme.TEXT_WARNING)
        self.text_area.tag_config("error", foreground=Theme.TEXT_RED, font=("Consolas", 9, "bold"))
        self.text_area.tag_config("accent", foreground=Theme.RED_GLOW, font=("Consolas", 9, "bold"))
        self.text_area.tag_config("dim", foreground=Theme.TEXT_DIM)

    def log(self, message: str, tag: str = "info"):
        self.text_area.config(state="normal")
        
        # Auto-detect status keywords for coloring
        lower = message.lower()
        if tag == "info":
            if "success" in lower or "completed" in lower or "saved to" in lower:
                tag = "success"
            elif "error" in lower or "failed" in lower or "exception" in lower:
                tag = "error"
            elif "warning" in lower:
                tag = "warning"
            elif "[agent]" in lower or "[user]" in lower or "extracting" in lower or "generating" in lower:
                tag = "accent"
            elif message.startswith("  "):
                tag = "dim"

        self.text_area.insert(tk.END, message + "\n", tag)
        self.text_area.see(tk.END)
        self.text_area.config(state="disabled")

    def clear(self):
        self.text_area.config(state="normal")
        self.text_area.delete("1.0", tk.END)
        self.text_area.config(state="disabled")


class StyledActionBtn(tk.Button):
    """Modern Hoverable Black/Red Button widget for key actions."""
    def __init__(
        self,
        parent,
        text: str,
        command: Optional[Callable] = None,
        is_primary: bool = True,
        is_danger: bool = False,
        **kwargs
    ):
        if is_danger or is_primary:
            bg_color = Theme.RED_PRIMARY
            hover_color = Theme.RED_HOVER
            active_color = Theme.RED_ACTIVE
            fg_color = Theme.TEXT_WHITE
        else:
            bg_color = Theme.BG_SURFACE
            hover_color = Theme.BG_SURFACE_HOVER
            active_color = Theme.BG_DARKEST
            fg_color = Theme.TEXT_WHITE

        self.normal_bg = bg_color
        self.hover_bg = hover_color
        self.active_bg = active_color

        padx = kwargs.pop("padx", 16)
        pady = kwargs.pop("pady", 8)
        font_btn = kwargs.pop("font", Theme.FONT_BUTTON)
        relief = kwargs.pop("relief", "flat")
        bd = kwargs.pop("bd", 0)

        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg_color,
            fg=fg_color,
            activebackground=active_color,
            activeforeground=Theme.TEXT_WHITE,
            font=font_btn,
            relief=relief,
            bd=bd,
            padx=padx,
            pady=pady,
            cursor="hand2",
            **kwargs
        )

        self.bind("<Enter>", lambda e: self._on_hover())
        self.bind("<Leave>", lambda e: self._on_leave())

    def _on_hover(self):
        if str(self["state"]) != "disabled":
            self.config(bg=self.hover_bg)

    def _on_leave(self):
        if str(self["state"]) != "disabled":
            self.config(bg=self.normal_bg)

    def set_state(self, state: str):
        self.config(state=state)
        if state == "disabled":
            self.config(bg=Theme.BG_SURFACE, fg=Theme.TEXT_DIM, cursor="arrow")
        else:
            self.config(bg=self.normal_bg, fg=Theme.TEXT_WHITE, cursor="hand2")
