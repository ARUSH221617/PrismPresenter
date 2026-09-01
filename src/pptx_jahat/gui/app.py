import os
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import json
import time
from typing import Optional, List, Dict, Any
from PIL import Image, ImageTk, ImageDraw

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR, COMPONENTS_DIR
from pptx_jahat.tools.pptx_engine import extract_all_templates, get_components_catalog
from pptx_jahat.tools.pptx_builder import build_pptx_with_agent, verify_and_auto_heal_pptx
from pptx_jahat.tools.preview import render_pptx_file_previews
from pptx_jahat.tools.template_analyzer import (
    analyze_template,
    analyze_all_templates,
    load_notes,
    save_notes,
    get_analyzed_templates,
    NOTE_FILE
)
from pptx_jahat.agent import AIAgent
from pptx_jahat.gui.components import (
    Theme,
    apply_theme_to_ttk,
    ModernCard,
    Badge,
    ConsoleLogWidget,
    StyledActionBtn
)


class PPTXJahatApp(tk.Tk):
    """
    Upgraded Modern Desktop GUI for PPTX Jahat with a Black & Crimson Red Theme.
    Features:
    - High contrast dark canvas with vibrant red focal points.
    - Custom styled header banner with status badges.
    - Advanced Tab Navigation:
      1. Slide Generator & Live Preview
      2. Deck & Template Manager (Manage generated PPTX & Reference Templates)
      3. Template Catalog & Components
      4. Autonomous AI Terminal
      5. Engine Settings (.env)
    - Interactive slide viewer with real-time scaling, slide indicator, and PowerPoint launcher.
    - Non-blocking asynchronous threading with live console stream tags.
    """

    def __init__(self):
        super().__init__()
        self.title("PPTX JAHAT — AI Presentation Generation Suite")
        self.geometry("1200x820")
        self.minsize(980, 680)
        self.configure(bg=Theme.BG_DARKEST)

        # Apply custom Black & Red TTK Theme
        apply_theme_to_ttk(self)

        self.current_generated_pptx: Optional[str] = None
        self.raw_preview_pil_images: List[Image.Image] = []
        self.preview_engine_name: str = ""
        self.preview_images_tk: List[ImageTk.PhotoImage] = []
        self.current_slide_idx: int = 0

        # Multi-tab Visual Previews State
        self.visual_latest_pil_images: List[Image.Image] = []
        self.visual_latest_idx: int = 0

        self.ai_test_pil_images: List[Dict[str, Any]] = []
        self.ai_test_idx: int = 0

        # Template Analyzer State (NEW)
        self.analyze_selected_file_path: Optional[str] = None
        self.analyze_preview_pil_images: List[Image.Image] = []
        self.analyze_preview_engine_name: str = ""
        self.analyze_current_slide_idx: int = 0

        # Manager state
        self.mgr_selected_file_path: Optional[str] = None
        self.mgr_preview_pil_images: List[Image.Image] = []
        self.mgr_preview_engine_name: str = ""
        self.mgr_current_slide_idx: int = 0

        self._build_header()
        self._build_tabs()
        self._build_status_bar()

    def _build_header(self):
        """Top branding header with high-tech black & red cyberpunk aesthetic."""
        header_frame = tk.Frame(self, bg=Theme.BG_HEADER, padx=20, pady=12)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        # Left branding
        brand_left = tk.Frame(header_frame, bg=Theme.BG_HEADER)
        brand_left.pack(side=tk.LEFT)

        # Red accent emblem / block
        emblem = tk.Label(
            brand_left,
            text="⚡ PPTX",
            bg=Theme.RED_PRIMARY,
            fg=Theme.TEXT_WHITE,
            font=("Segoe UI", 11, "bold"),
            padx=8,
            pady=2
        )
        emblem.pack(side=tk.LEFT, padx=(0, 10))

        title_box = tk.Frame(brand_left, bg=Theme.BG_HEADER)
        title_box.pack(side=tk.LEFT)

        app_title = tk.Label(
            title_box,
            text="JAHAT PRESENTATION AGENT",
            bg=Theme.BG_HEADER,
            fg=Theme.TEXT_WHITE,
            font=("Segoe UI", 13, "bold"),
            anchor="w"
        )
        app_title.pack(anchor="w")

        app_subtitle = tk.Label(
            title_box,
            text="Autonomous Slide Synthesizer • Dynamic Templates • 9Router Engine",
            bg=Theme.BG_HEADER,
            fg=Theme.TEXT_MUTED,
            font=("Segoe UI", 8),
            anchor="w"
        )
        app_subtitle.pack(anchor="w")

        # Right badges
        badges_frame = tk.Frame(header_frame, bg=Theme.BG_HEADER)
        badges_frame.pack(side=tk.RIGHT)

        self.badge_status = Badge(badges_frame, text="● SYSTEM READY", bg_color=Theme.BADGE_BG_RED, fg_color=Theme.BADGE_TEXT_RED)
        self.badge_status.pack(side=tk.RIGHT, padx=4)

        self.badge_model = Badge(
            badges_frame,
            text=f"MODEL: {Config.NINEROUTER_CHAT_MODEL.split('/')[-1]}",
            bg_color=Theme.BG_SURFACE,
            fg_color=Theme.TEXT_MAIN,
            border_color=Theme.BORDER_DARK
        )
        self.badge_model.pack(side=tk.RIGHT, padx=4)

        # Divider line
        sep = tk.Frame(self, bg=Theme.BORDER_DARK, height=1)
        sep.pack(fill=tk.X, side=tk.TOP)

    def _build_tabs(self):
        """Constructs modern styled notebook tabs."""
        notebook_container = tk.Frame(self, bg=Theme.BG_DARKEST, padx=12, pady=10)
        notebook_container.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(notebook_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Docx -> PPTX Generator & In-App Preview
        self.tab_generator = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_generator, text="  ⚡ Slide Generator & Live Preview  ")
        self._setup_generator_tab()

        # Tab 2: Template Intelligence & AI Analyzer (NEW)
        self.tab_analyze = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_analyze, text="  🔍 Template Intelligence & Analyze  ")
        self._setup_analyze_tab()

        # Tab 3: Deck & Template Manager
        self.tab_manager = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_manager, text="  🗂️ Deck & Template Manager  ")
        self._setup_manager_tab()

        # Tab 3: Template & Components Catalog
        self.tab_components = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_components, text="  📁 Template Catalog & Components  ")
        self._setup_components_tab()

        # Tab 4: AI Agent Chat & Automation
        self.tab_agent = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_agent, text="  🤖 Autonomous AI Terminal  ")
        self._setup_agent_tab()

        # Tab 5: Settings & Configuration
        self.tab_settings = tk.Frame(self.notebook, bg=Theme.BG_MAIN, padx=10, pady=10)
        self.notebook.add(self.tab_settings, text="  ⚙️ Engine Settings (.env)  ")
        self._setup_settings_tab()

    def _build_status_bar(self):
        """Bottom dark status bar."""
        bar = tk.Frame(self, bg=Theme.BG_HEADER, padx=12, pady=5)
        bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_left_var = tk.StringVar(value="Idle • Ready for presentation generation")
        lbl_status = tk.Label(
            bar,
            textvariable=self.status_left_var,
            bg=Theme.BG_HEADER,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_CAPTION
        )
        lbl_status.pack(side=tk.LEFT)

        lbl_version = tk.Label(
            bar,
            text="PPTX Jahat v1.0.0 • Black & Crimson UI",
            bg=Theme.BG_HEADER,
            fg=Theme.TEXT_DIM,
            font=Theme.FONT_CAPTION
        )
        lbl_version.pack(side=tk.RIGHT)

        sep = tk.Frame(self, bg=Theme.BORDER_DARK, height=1)
        sep.pack(fill=tk.X, side=tk.BOTTOM)

    # -------------------------------------------------------------
    # TAB 1: Generator & Interactive Slide Preview
    # -------------------------------------------------------------
    def _setup_generator_tab(self):
        paned = tk.PanedWindow(
            self.tab_generator,
            orient=tk.HORIZONTAL,
            bg=Theme.BG_DARKEST,
            bd=0,
            sashwidth=4,
            sashrelief="flat"
        )
        paned.pack(fill=tk.BOTH, expand=True)

        # --- LEFT PANEL: Settings, Inputs & Generation Terminal ---
        left_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(left_container, minsize=420, stretch="always")

        # Generator Card
        gen_card = ModernCard(
            left_container,
            title="SLIDE SYNTHESIZER",
            subtitle="Transform Word Docx into Professional PPTX",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        gen_card.pack(fill=tk.X, pady=(0, 10))

        # Row 1: Word Document Input
        r1 = tk.Frame(gen_card.body, bg=Theme.BG_SURFACE)
        r1.pack(fill=tk.X, pady=4)
        tk.Label(r1, text="Word Document (.docx):", bg=Theme.BG_SURFACE, fg=Theme.TEXT_MAIN, font=Theme.FONT_BODY_BOLD, width=20, anchor="w").pack(side=tk.LEFT)
        self.docx_path_var = tk.StringVar()
        entry_docx = ttk.Entry(r1, textvariable=self.docx_path_var)
        entry_docx.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        btn_b_docx = ttk.Button(r1, text="Browse", style="Secondary.TButton", command=self._browse_docx)
        btn_b_docx.pack(side=tk.LEFT)

        # Row 2: Template Selection
        r2 = tk.Frame(gen_card.body, bg=Theme.BG_SURFACE)
        r2.pack(fill=tk.X, pady=4)
        tk.Label(r2, text="Base Template Style:", bg=Theme.BG_SURFACE, fg=Theme.TEXT_MAIN, font=Theme.FONT_BODY_BOLD, width=20, anchor="w").pack(side=tk.LEFT)
        self.template_choice_var = tk.StringVar()
        self.template_combo = ttk.Combobox(r2, textvariable=self.template_choice_var, state="readonly")
        self.template_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        btn_refresh = ttk.Button(r2, text="Refresh", style="Secondary.TButton", command=self._refresh_templates)
        btn_refresh.pack(side=tk.LEFT)
        btn_jump_analyze = ttk.Button(
            r2,
            text="🔍 Intelligence Notes",
            style="Secondary.TButton",
            command=lambda: self.notebook.select(self.tab_analyze)
        )
        btn_jump_analyze.pack(side=tk.LEFT, padx=(4, 0))

        # Row 3: Output PPTX Destination
        r3 = tk.Frame(gen_card.body, bg=Theme.BG_SURFACE)
        r3.pack(fill=tk.X, pady=4)
        tk.Label(r3, text="Target Output (.pptx):", bg=Theme.BG_SURFACE, fg=Theme.TEXT_MAIN, font=Theme.FONT_BODY_BOLD, width=20, anchor="w").pack(side=tk.LEFT)
        self.output_path_var = tk.StringVar()
        entry_out = ttk.Entry(r3, textvariable=self.output_path_var)
        entry_out.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        btn_b_out = ttk.Button(r3, text="Save As", style="Secondary.TButton", command=self._browse_output)
        btn_b_out.pack(side=tk.LEFT)

        # Action Buttons Row
        action_bar = tk.Frame(gen_card.body, bg=Theme.BG_SURFACE)
        action_bar.pack(fill=tk.X, pady=(12, 4))

        self.btn_generate = StyledActionBtn(
            action_bar,
            text="⚡ Generate Presentation",
            command=self._run_generator,
            is_primary=True
        )
        self.btn_generate.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_pptx = StyledActionBtn(
            action_bar,
            text="📊 Open in PowerPoint",
            command=self._open_in_powerpoint,
            is_primary=False
        )
        self.btn_open_pptx.set_state("disabled")
        self.btn_open_pptx.pack(side=tk.LEFT)

        # Console Logs Card
        self.gen_console = ConsoleLogWidget(left_container, title="Generation Execution Stream", height=10)
        self.gen_console.pack(fill=tk.BOTH, expand=True)

        # --- RIGHT PANEL: Interactive In-App Multi-Tab Preview ---
        right_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(right_container, minsize=480, stretch="always")

        preview_card = ModernCard(
            right_container,
            title="VISUAL DECK & AI INSPECTOR",
            subtitle="Live Generated Slides, High-Res Visuals & AI Multimodal Payloads",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        preview_card.pack(fill=tk.BOTH, expand=True)

        # Multi-tab sub notebook inside generator preview card
        self.preview_sub_notebook = ttk.Notebook(preview_card.body)
        self.preview_sub_notebook.pack(fill=tk.BOTH, expand=True)

        # SUB-TAB 1: Preview (Interactive Slide Viewer)
        self.subtab_preview = tk.Frame(self.preview_sub_notebook, bg=Theme.BG_SURFACE)
        self.preview_sub_notebook.add(self.subtab_preview, text="  📊 Live Deck Preview  ")
        self._setup_subtab_deck_preview()

        # SUB-TAB 2: Visual Image (Screenshot of Latest Generated Slides)
        self.subtab_visual = tk.Frame(self.preview_sub_notebook, bg=Theme.BG_SURFACE)
        self.preview_sub_notebook.add(self.subtab_visual, text="  🖼️ Visual Screenshots  ")
        self._setup_subtab_visual_images()

        # SUB-TAB 3: Visual AI Test Image (Images Sent to AI Agent)
        self.subtab_ai_test = tk.Frame(self.preview_sub_notebook, bg=Theme.BG_SURFACE)
        self.preview_sub_notebook.add(self.subtab_ai_test, text="  🤖 Visual AI Test Images  ")
        self._setup_subtab_ai_test_images()

        self._refresh_templates()

    def _setup_subtab_deck_preview(self):
        # Nav & Controls Toolbar
        nav_toolbar = tk.Frame(self.subtab_preview, bg=Theme.BG_SURFACE, pady=4)
        nav_toolbar.pack(fill=tk.X, side=tk.TOP)

        self.btn_prev_slide = StyledActionBtn(
            nav_toolbar,
            text="◀ Prev Slide",
            command=self._prev_slide,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_prev_slide.set_state("disabled")
        self.btn_prev_slide.pack(side=tk.LEFT, padx=4)

        self.slide_counter_var = tk.StringVar(value="No slides loaded")
        self.preview_engine_badge_var = tk.StringVar(value="")

        center_info_box = tk.Frame(nav_toolbar, bg=Theme.BG_SURFACE)
        center_info_box.pack(side=tk.LEFT, expand=True)

        lbl_counter = tk.Label(
            center_info_box,
            textvariable=self.slide_counter_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_counter.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_preview_engine = tk.Label(
            center_info_box,
            textvariable=self.preview_engine_badge_var,
            bg=Theme.BADGE_BG_RED,
            fg=Theme.BADGE_TEXT_RED,
            font=Theme.FONT_CAPTION,
            padx=6,
            pady=1,
            highlightbackground=Theme.BADGE_BORDER_RED,
            highlightthickness=1
        )
        self.lbl_preview_engine.pack(side=tk.LEFT)

        self.btn_next_slide = StyledActionBtn(
            nav_toolbar,
            text="Next Slide ▶",
            command=self._next_slide,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_next_slide.set_state("disabled")
        self.btn_next_slide.pack(side=tk.RIGHT, padx=4)

        # Slide Display Canvas Box
        self.preview_display_box = tk.Frame(
            self.subtab_preview,
            bg=Theme.BG_DARKEST,
            highlightbackground=Theme.BORDER_DARK,
            highlightthickness=1
        )
        self.preview_display_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.preview_display_box.bind("<Configure>", lambda e: self._on_preview_resize())

        self.preview_label = tk.Label(
            self.preview_display_box,
            text="Generated slides and layouts will appear here automatically.",
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_BODY,
            anchor="center"
        )
        self.preview_label.pack(fill=tk.BOTH, expand=True)

    def _setup_subtab_visual_images(self):
        # Nav & Controls Toolbar for Visual Screenshots
        nav_toolbar = tk.Frame(self.subtab_visual, bg=Theme.BG_SURFACE, pady=4)
        nav_toolbar.pack(fill=tk.X, side=tk.TOP)

        self.btn_prev_visual = StyledActionBtn(
            nav_toolbar,
            text="◀ Prev Image",
            command=self._prev_visual_image,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_prev_visual.set_state("disabled")
        self.btn_prev_visual.pack(side=tk.LEFT, padx=4)

        self.visual_counter_var = tk.StringVar(value="No visual screenshot loaded")
        lbl_counter = tk.Label(
            nav_toolbar,
            textvariable=self.visual_counter_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_counter.pack(side=tk.LEFT, expand=True)

        self.btn_next_visual = StyledActionBtn(
            nav_toolbar,
            text="Next Image ▶",
            command=self._next_visual_image,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_next_visual.set_state("disabled")
        self.btn_next_visual.pack(side=tk.RIGHT, padx=4)

        # Slide Display Canvas Box
        self.visual_display_box = tk.Frame(
            self.subtab_visual,
            bg=Theme.BG_DARKEST,
            highlightbackground=Theme.BORDER_DARK,
            highlightthickness=1
        )
        self.visual_display_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.visual_display_box.bind("<Configure>", lambda e: self._on_visual_resize())

        self.visual_label = tk.Label(
            self.visual_display_box,
            text="Rendered high-res visual screenshots of latest slides will appear here.",
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_BODY,
            anchor="center"
        )
        self.visual_label.pack(fill=tk.BOTH, expand=True)

    def _setup_subtab_ai_test_images(self):
        # Nav & Controls Toolbar for AI Test Images
        nav_toolbar = tk.Frame(self.subtab_ai_test, bg=Theme.BG_SURFACE, pady=4)
        nav_toolbar.pack(fill=tk.X, side=tk.TOP)

        self.btn_prev_ai_test = StyledActionBtn(
            nav_toolbar,
            text="◀ Prev AI Image",
            command=self._prev_ai_test_image,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_prev_ai_test.set_state("disabled")
        self.btn_prev_ai_test.pack(side=tk.LEFT, padx=4)

        self.ai_test_counter_var = tk.StringVar(value="No AI test payload loaded")
        lbl_counter = tk.Label(
            nav_toolbar,
            textvariable=self.ai_test_counter_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_counter.pack(side=tk.LEFT, expand=True)

        self.btn_next_ai_test = StyledActionBtn(
            nav_toolbar,
            text="Next AI Image ▶",
            command=self._next_ai_test_image,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.btn_next_ai_test.set_state("disabled")
        self.btn_next_ai_test.pack(side=tk.RIGHT, padx=4)

        # Slide Display Canvas Box
        self.ai_test_display_box = tk.Frame(
            self.subtab_ai_test,
            bg=Theme.BG_DARKEST,
            highlightbackground=Theme.BORDER_DARK,
            highlightthickness=1
        )
        self.ai_test_display_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.ai_test_display_box.bind("<Configure>", lambda e: self._on_ai_test_resize())

        self.ai_test_label = tk.Label(
            self.ai_test_display_box,
            text="Visual screenshots sent to 9Router Vision AI Agent will be displayed here.",
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_BODY,
            anchor="center"
        )
        self.ai_test_label.pack(fill=tk.BOTH, expand=True)

    def _browse_docx(self):
        f = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
        if f:
            self.docx_path_var.set(f)
            if not self.output_path_var.get():
                p = Path(f)
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                self.output_path_var.set(str(OUTPUT_DIR / f"{p.stem}_generated.pptx"))

    def _browse_output(self):
        f = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint Presentation", "*.pptx")])
        if f:
            self.output_path_var.set(f)

    def _refresh_templates(self):
        pptxs = [p.name for p in DATA_DIR.glob("*.pptx") if not p.name.endswith("_generated.pptx")]
        self.template_combo["values"] = ["All Templates (Global AI Matching)"] + pptxs
        if pptxs:
            self.template_combo.current(0)
        else:
            self.template_choice_var.set("No templates found in data/")

    def _open_in_powerpoint(self):
        if not self.current_generated_pptx or not Path(self.current_generated_pptx).exists():
            messagebox.showwarning("Warning", "No generated PPTX file available to open.")
            return
            
        self._launch_file(self.current_generated_pptx)

    def _launch_file(self, file_path: str):
        try:
            os.startfile(file_path)
            self.gen_console.log(f"Opened in default app: {file_path}", "success")
        except Exception:
            try:
                subprocess.Popen(["start", "", file_path], shell=True)
                self.gen_console.log(f"Launched file: {file_path}", "info")
            except Exception as ex:
                messagebox.showerror("Error", f"Could not open file: {str(ex)}")

    def _reveal_in_explorer(self, file_path: str):
        p = Path(file_path).resolve()
        if not p.exists():
            messagebox.showwarning("Warning", f"File not found: {file_path}")
            return
        try:
            subprocess.Popen(f'explorer /select,"{str(p)}"')
        except Exception as e:
            messagebox.showerror("Error", f"Could not reveal in Explorer: {str(e)}")

    def _on_preview_resize(self):
        """Re-render current slide if window dimensions change."""
        if self.raw_preview_pil_images and self.current_slide_idx < len(self.raw_preview_pil_images):
            self._update_preview_display()

    def _on_visual_resize(self):
        """Re-render visual screenshot if window dimensions change."""
        if self.visual_latest_pil_images and self.visual_latest_idx < len(self.visual_latest_pil_images):
            self._update_visual_display()

    def _on_ai_test_resize(self):
        """Re-render AI test image if window dimensions change."""
        if self.ai_test_pil_images and self.ai_test_idx < len(self.ai_test_pil_images):
            self._update_ai_test_display()

    def _load_slide_previews(self, pptx_path: str):
        try:
            self.gen_console.log("Rendering high-res slide visual previews...", "accent")
            res = render_pptx_file_previews(pptx_path, target_width_px=800, return_engine_info=True)
            if isinstance(res, tuple):
                self.raw_preview_pil_images, self.preview_engine_name = res
            else:
                self.raw_preview_pil_images = res
                self.preview_engine_name = "Renderer"
            self.current_slide_idx = 0
            self._update_preview_display()

            # Update Subtab 2: Visual screenshots of latest generated slides
            self.visual_latest_pil_images = list(self.raw_preview_pil_images)
            self.visual_latest_idx = 0
            self._update_visual_display()

            self.gen_console.log(f"Successfully rendered {len(self.raw_preview_pil_images)} slides via {self.preview_engine_name}.", "success")
            if hasattr(self, "_refresh_manager_lists"):
                self._refresh_manager_lists()
        except Exception as e:
            self.gen_console.log(f"Preview render warning: {str(e)}", "warning")

    def _set_ai_test_images(self, sent_images: List[Dict[str, Any]]):
        """Receives base64 template slide snapshots sent directly to the AI agent."""
        parsed_images = []
        for item in sent_images:
            b64 = item.get("base64")
            if b64 and "," in b64:
                try:
                    import base64
                    import io
                    raw_b64 = b64.split(",", 1)[1]
                    img_data = base64.b64decode(raw_b64)
                    pil_img = Image.open(io.BytesIO(img_data)).convert("RGB")
                    parsed_images.append({
                        "image": pil_img,
                        "template_file": item.get("template_file", "Template"),
                        "slide_index": item.get("slide_index", 0),
                        "archetype": item.get("archetype", "Archetype")
                    })
                except Exception:
                    pass
        self.ai_test_pil_images = parsed_images
        self.ai_test_idx = 0
        self._update_ai_test_display()

    def _update_preview_display(self):
        if not self.raw_preview_pil_images:
            self.preview_label.config(image="", text="No preview available.")
            self.slide_counter_var.set("0 / 0")
            self.preview_engine_badge_var.set("")
            self.lbl_preview_engine.config(bg=Theme.BG_SURFACE, highlightthickness=0)
            self.btn_prev_slide.set_state("disabled")
            self.btn_next_slide.set_state("disabled")
            return

        total = len(self.raw_preview_pil_images)
        self.slide_counter_var.set(f"Slide {self.current_slide_idx + 1} of {total}")
        
        # Engine badge indicator
        if self.preview_engine_name:
            if "PowerPoint" in self.preview_engine_name:
                self.preview_engine_badge_var.set("⚡ Native PowerPoint")
                self.lbl_preview_engine.config(bg="#1e3a29", fg="#4ade80", highlightbackground="#22c55e", highlightthickness=1)
            else:
                self.preview_engine_badge_var.set("🎨 Pure PIL Engine")
                self.lbl_preview_engine.config(bg=Theme.BADGE_BG_RED, fg=Theme.BADGE_TEXT_RED, highlightbackground=Theme.BADGE_BORDER_RED, highlightthickness=1)
        else:
            self.preview_engine_badge_var.set("")
            self.lbl_preview_engine.config(bg=Theme.BG_SURFACE, highlightthickness=0)
        
        # Calculate fit within preview container
        box_w = max(100, self.preview_display_box.winfo_width() - 20)
        box_h = max(100, self.preview_display_box.winfo_height() - 20)

        raw_img = self.raw_preview_pil_images[self.current_slide_idx]
        img_w, img_h = raw_img.size

        scale = min(box_w / img_w, box_h / img_h, 1.0)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))

        resized = raw_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        # The preview is RGBA; if it has transparent pixels (e.g. from gradients
        # with alpha, noFill backgrounds, or PNG master images), the dark Tk
        # frame (#0b0c10) bleeds through and makes the slide look like a wireframe.
        # Composite onto a solid white background so the preview always looks like
        # the final rendered slide as it would appear in PowerPoint.
        if resized.mode == "RGBA":
            bg = Image.new("RGBA", resized.size, (255, 255, 255, 255))
            bg.alpha_composite(resized)
            resized = bg.convert("RGB")
        self._current_tk_img = ImageTk.PhotoImage(resized)
        
        self.preview_label.config(image=self._current_tk_img, text="")
        
        self.btn_prev_slide.set_state("normal" if self.current_slide_idx > 0 else "disabled")
        self.btn_next_slide.set_state("normal" if self.current_slide_idx < total - 1 else "disabled")

    def _prev_slide(self):
        if self.current_slide_idx > 0:
            self.current_slide_idx -= 1
            self._update_preview_display()

    def _next_slide(self):
        if self.current_slide_idx < len(self.raw_preview_pil_images) - 1:
            self.current_slide_idx += 1
            self._update_preview_display()

    def _update_visual_display(self):
        if not self.visual_latest_pil_images:
            self.visual_label.config(image="", text="No visual screenshots available.")
            self.visual_counter_var.set("0 / 0")
            self.btn_prev_visual.set_state("disabled")
            self.btn_next_visual.set_state("disabled")
            return

        total = len(self.visual_latest_pil_images)
        self.visual_counter_var.set(f"Visual Screenshot {self.visual_latest_idx + 1} of {total}")

        box_w = max(100, self.visual_display_box.winfo_width() - 20)
        box_h = max(100, self.visual_display_box.winfo_height() - 20)

        raw_img = self.visual_latest_pil_images[self.visual_latest_idx]
        img_w, img_h = raw_img.size

        scale = min(box_w / img_w, box_h / img_h, 1.0)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))

        resized = raw_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        if resized.mode == "RGBA":
            bg = Image.new("RGBA", resized.size, (255, 255, 255, 255))
            bg.alpha_composite(resized)
            resized = bg.convert("RGB")
        self._visual_current_tk_img = ImageTk.PhotoImage(resized)

        self.visual_label.config(image=self._visual_current_tk_img, text="")

        self.btn_prev_visual.set_state("normal" if self.visual_latest_idx > 0 else "disabled")
        self.btn_next_visual.set_state("normal" if self.visual_latest_idx < total - 1 else "disabled")

    def _prev_visual_image(self):
        if self.visual_latest_idx > 0:
            self.visual_latest_idx -= 1
            self._update_visual_display()

    def _next_visual_image(self):
        if self.visual_latest_idx < len(self.visual_latest_pil_images) - 1:
            self.visual_latest_idx += 1
            self._update_visual_display()

    def _update_ai_test_display(self):
        if not self.ai_test_pil_images:
            self.ai_test_label.config(image="", text="No AI test images sent yet.")
            self.ai_test_counter_var.set("0 / 0")
            self.btn_prev_ai_test.set_state("disabled")
            self.btn_next_ai_test.set_state("disabled")
            return

        total = len(self.ai_test_pil_images)
        entry = self.ai_test_pil_images[self.ai_test_idx]
        tpl_name = entry.get("template_file", "")
        s_idx = entry.get("slide_index", 0)
        arch = entry.get("archetype", "")
        self.ai_test_counter_var.set(f"AI Payload {self.ai_test_idx + 1}/{total} • {tpl_name} [Slide {s_idx+1}] ({arch})")

        box_w = max(100, self.ai_test_display_box.winfo_width() - 20)
        box_h = max(100, self.ai_test_display_box.winfo_height() - 20)

        raw_img = entry["image"]
        img_w, img_h = raw_img.size

        scale = min(box_w / img_w, box_h / img_h, 1.0)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))

        resized = raw_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        if resized.mode == "RGBA":
            bg = Image.new("RGBA", resized.size, (255, 255, 255, 255))
            bg.alpha_composite(resized)
            resized = bg.convert("RGB")
        self._ai_test_current_tk_img = ImageTk.PhotoImage(resized)

        self.ai_test_label.config(image=self._ai_test_current_tk_img, text="")

        self.btn_prev_ai_test.set_state("normal" if self.ai_test_idx > 0 else "disabled")
        self.btn_next_ai_test.set_state("normal" if self.ai_test_idx < total - 1 else "disabled")

    def _prev_ai_test_image(self):
        if self.ai_test_idx > 0:
            self.ai_test_idx -= 1
            self._update_ai_test_display()

    def _next_ai_test_image(self):
        if self.ai_test_idx < len(self.ai_test_pil_images) - 1:
            self.ai_test_idx += 1
            self._update_ai_test_display()

    def _run_generator(self):
        docx_p = self.docx_path_var.get().strip()
        if not docx_p or not Path(docx_p).exists():
            messagebox.showerror("Error", "Please select a valid Word (.docx) document.")
            return

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_p = self.output_path_var.get().strip() or str(OUTPUT_DIR / "generated_presentation.pptx")
        tpl = self.template_choice_var.get()
        tpl_name = None if "All Templates" in tpl or "No templates" in tpl else tpl

        self.btn_generate.set_state("disabled")
        self.btn_open_pptx.set_state("disabled")
        self.badge_status.set_text("● GENERATING...", fg_color=Theme.TEXT_WHITE, bg_color=Theme.RED_PRIMARY)
        self.status_left_var.set("Synthesizing slides from Word document...")

        def worker():
            def log_fn(msg):
                self.gen_console.log(msg)

            def ai_images_cb(sent_images):
                self.after(50, lambda: self._set_ai_test_images(sent_images))

            try:
                self.gen_console.log(f"Starting PPTX generation pipeline for '{Path(docx_p).name}'", "accent")
                res = build_pptx_with_agent(
                    docx_p,
                    out_p,
                    tpl_name,
                    log_callback=log_fn,
                    on_ai_images_ready=ai_images_cb
                )
                self.current_generated_pptx = res
                self.gen_console.log(f"SUCCESS: Generated PPTX saved at {res}", "success")
                
                # Render previews inside Tkinter GUI
                self.after(100, lambda: self._load_slide_previews(res))
                self.after(100, lambda: self.btn_open_pptx.set_state("normal"))
                self.after(100, lambda: self.badge_status.set_text("● BUILD READY", fg_color=Theme.TEXT_SUCCESS, bg_color=Theme.BG_SURFACE))
                self.after(100, lambda: self.status_left_var.set(f"Completed: {Path(res).name}"))
                
                messagebox.showinfo("Success", f"Presentation created successfully!\nSaved to: {res}")
            except Exception as e:
                self.gen_console.log(f"ERROR: {str(e)}", "error")
                self.after(100, lambda: self.badge_status.set_text("● BUILD FAILED", fg_color=Theme.TEXT_RED, bg_color=Theme.BADGE_BG_RED))
                self.after(100, lambda: self.status_left_var.set("Generation error occurred."))
                messagebox.showerror("Generation Error", str(e))
            finally:
                self.after(100, lambda: self.btn_generate.set_state("normal"))

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------------------
    # TAB 2: Template Intelligence & AI Analyzer (Analyze templates -> NOTE.md)
    # -------------------------------------------------------------
    def _setup_analyze_tab(self):
        paned = tk.PanedWindow(
            self.tab_analyze,
            orient=tk.HORIZONTAL,
            bg=Theme.BG_DARKEST,
            bd=0,
            sashwidth=4,
            sashrelief="flat"
        )
        paned.pack(fill=tk.BOTH, expand=True)

        # --- LEFT PANEL: Template List & Live Archetype/Slide Inspector ---
        left_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(left_container, minsize=500, stretch="always")

        # Top Control Card for Template Analysis
        top_ctrl_card = ModernCard(
            left_container,
            title="TEMPLATE REPOSITORY & AI ANALYSIS",
            subtitle="Analyze PPTX designs with AI Agent & save style notes to data/NOTE.md",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        top_ctrl_card.pack(fill=tk.X, pady=(0, 8))

        btn_bar = tk.Frame(top_ctrl_card.body, bg=Theme.BG_SURFACE)
        btn_bar.pack(fill=tk.X, pady=2)

        self.btn_analyze_sel = StyledActionBtn(
            btn_bar,
            text="⚡ Analyze Selected Template",
            command=self._analyze_selected_template,
            is_primary=True,
            padx=10,
            pady=5
        )
        self.btn_analyze_sel.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_analyze_all = StyledActionBtn(
            btn_bar,
            text="🚀 Analyze All Templates (Batch)",
            command=self._analyze_all_templates_batch,
            is_primary=False,
            padx=10,
            pady=5
        )
        self.btn_analyze_all.pack(side=tk.LEFT, padx=4)

        StyledActionBtn(
            btn_bar,
            text="🔄 Refresh List",
            command=self._refresh_analyze_templates_list,
            is_primary=False,
            padx=10,
            pady=5
        ).pack(side=tk.RIGHT)

        # Progress bar frame
        prog_frame = tk.Frame(top_ctrl_card.body, bg=Theme.BG_SURFACE)
        prog_frame.pack(fill=tk.X, pady=(6, 0))

        self.analyze_progress_var = tk.DoubleVar(value=0.0)
        self.analyze_progress_bar = ttk.Progressbar(
            prog_frame,
            variable=self.analyze_progress_var,
            maximum=100,
            mode="determinate"
        )
        self.analyze_progress_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 2))

        self.analyze_progress_status_var = tk.StringVar(value="Ready to analyze templates")
        lbl_prog_status = tk.Label(
            prog_frame,
            textvariable=self.analyze_progress_status_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_CAPTION,
            anchor="w"
        )
        lbl_prog_status.pack(fill=tk.X, side=tk.LEFT)

        # Templates Treeview Card
        tree_card = ModernCard(
            left_container,
            title="AVAILABLE PPTX TEMPLATES (data/*.pptx)",
            subtitle="Select a template to view archetype metadata, slide preview, or run AI analysis",
            accent_color=Theme.RED_PRIMARY
        )
        tree_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        tree_frame = tk.Frame(tree_card.body, bg=Theme.BG_DARKEST)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("slides", "status", "style", "purpose")
        self.tree_analyze_templates = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="tree headings",
            selectmode="browse",
            height=6
        )
        self.tree_analyze_templates.heading("#0", text="Template Name", anchor="w")
        self.tree_analyze_templates.heading("slides", text="Slides", anchor="center")
        self.tree_analyze_templates.heading("status", text="NOTE.md Status", anchor="center")
        self.tree_analyze_templates.heading("style", text="Style / Mood", anchor="w")
        self.tree_analyze_templates.heading("purpose", text="Target Purpose", anchor="w")

        self.tree_analyze_templates.column("#0", width=120, minwidth=100)
        self.tree_analyze_templates.column("slides", width=60, minwidth=50, anchor="center")
        self.tree_analyze_templates.column("status", width=110, minwidth=90, anchor="center")
        self.tree_analyze_templates.column("style", width=110, minwidth=80)
        self.tree_analyze_templates.column("purpose", width=140, minwidth=100)

        # Configure tag colors
        self.tree_analyze_templates.tag_configure("analyzed", foreground=Theme.TEXT_SUCCESS)
        self.tree_analyze_templates.tag_configure("pending", foreground=Theme.TEXT_MUTED)

        sb_tree_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_analyze_templates.yview)
        self.tree_analyze_templates.configure(yscrollcommand=sb_tree_y.set)
        sb_tree_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_analyze_templates.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_analyze_templates.bind("<<TreeviewSelect>>", self._on_analyze_tree_selected)

        # Bottom Sub-frame: Selected Template Metadata + Live Slide Preview
        bottom_box = tk.Frame(left_container, bg=Theme.BG_MAIN)
        bottom_box.pack(fill=tk.BOTH, expand=True)

        # Metadata Card
        self.analyze_meta_card = ModernCard(
            bottom_box,
            title="TEMPLATE DETAILS & ARCHETYPES",
            subtitle="Select a template above to inspect structure",
            accent_color=Theme.RED_PRIMARY
        )
        self.analyze_meta_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        self.analyze_meta_text = tk.Text(
            self.analyze_meta_card.body,
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MAIN,
            font=Theme.FONT_CAPTION,
            wrap=tk.WORD,
            height=8,
            bd=0,
            padx=6,
            pady=6
        )
        self.analyze_meta_text.pack(fill=tk.BOTH, expand=True)
        self.analyze_meta_text.config(state="disabled")

        # Slide Visual Preview Card
        self.analyze_preview_card = ModernCard(
            bottom_box,
            title="SLIDE VISUAL PREVIEW",
            subtitle="Rendered template slide preview",
            accent_color=Theme.RED_PRIMARY
        )
        self.analyze_preview_card.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(4, 0))

        # Preview navigation toolbar
        prev_nav_bar = tk.Frame(self.analyze_preview_card.body, bg=Theme.BG_SURFACE)
        prev_nav_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 4))

        self.btn_analyze_prev = StyledActionBtn(
            prev_nav_bar,
            text="◀ Prev",
            command=self._analyze_prev_slide,
            is_primary=False,
            padx=6,
            pady=2
        )
        self.btn_analyze_prev.pack(side=tk.LEFT)

        self.analyze_slide_counter_var = tk.StringVar(value="0 / 0")
        lbl_slide_cnt = tk.Label(
            prev_nav_bar,
            textvariable=self.analyze_slide_counter_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_WHITE,
            font=Theme.FONT_BODY_BOLD
        )
        lbl_slide_cnt.pack(side=tk.LEFT, padx=6)

        self.btn_analyze_next = StyledActionBtn(
            prev_nav_bar,
            text="Next ▶",
            command=self._analyze_next_slide,
            is_primary=False,
            padx=6,
            pady=2
        )
        self.btn_analyze_next.pack(side=tk.LEFT)

        StyledActionBtn(
            prev_nav_bar,
            text="📊 Open PPT",
            command=self._analyze_open_selected_in_ppt,
            is_primary=False,
            padx=6,
            pady=2
        ).pack(side=tk.RIGHT)

        self.analyze_preview_display_box = tk.Frame(self.analyze_preview_card.body, bg=Theme.BG_DARKEST)
        self.analyze_preview_display_box.pack(fill=tk.BOTH, expand=True)

        self.analyze_preview_label = tk.Label(
            self.analyze_preview_display_box,
            text="Select a template to view slides",
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_BODY
        )
        self.analyze_preview_label.pack(fill=tk.BOTH, expand=True)
        self.analyze_preview_display_box.bind("<Configure>", lambda e: self._on_analyze_preview_resize())

        # --- RIGHT PANEL: Rich Markdown Notes (data/NOTE.md) & Live Execution Console ---
        right_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(right_container, minsize=520, stretch="always")

        notes_card = ModernCard(
            right_container,
            title="TEMPLATE INTELLIGENCE NOTES (data/NOTE.md)",
            subtitle="AI knowledge base: Purpose, Content Brief, Visual Ideas, Style & Archetype matching",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        notes_card.pack(fill=tk.BOTH, expand=True)

        # Top Action Toolbar for Notes
        notes_toolbar = tk.Frame(notes_card.body, bg=Theme.BG_SURFACE)
        notes_toolbar.pack(fill=tk.X, side=tk.TOP, pady=(0, 6))

        StyledActionBtn(
            notes_toolbar,
            text="💾 Save NOTE.md",
            command=self._save_analyze_notes,
            is_primary=True,
            padx=10,
            pady=4
        ).pack(side=tk.LEFT, padx=(0, 6))

        StyledActionBtn(
            notes_toolbar,
            text="🔄 Reload from Disk",
            command=self._reload_analyze_notes,
            is_primary=False,
            padx=10,
            pady=4
        ).pack(side=tk.LEFT, padx=4)

        StyledActionBtn(
            notes_toolbar,
            text="📋 Copy Notes",
            command=self._copy_analyze_notes,
            is_primary=False,
            padx=10,
            pady=4
        ).pack(side=tk.LEFT, padx=4)

        StyledActionBtn(
            notes_toolbar,
            text="📁 Open Folder",
            command=lambda: self._reveal_in_explorer(str(DATA_DIR)),
            is_primary=False,
            padx=10,
            pady=4
        ).pack(side=tk.LEFT, padx=4)

        self.notes_status_var = tk.StringVar(value="File: data/NOTE.md")
        lbl_notes_status = tk.Label(
            notes_toolbar,
            textvariable=self.notes_status_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_notes_status.pack(side=tk.RIGHT, padx=6)

        # Split pane between Notes Editor & Console Terminal
        right_split = tk.PanedWindow(
            notes_card.body,
            orient=tk.VERTICAL,
            bg=Theme.BG_DARKEST,
            bd=0,
            sashwidth=4,
            sashrelief="flat"
        )
        right_split.pack(fill=tk.BOTH, expand=True)

        # Upper: Notes Editor Frame
        editor_frame = tk.Frame(right_split, bg=Theme.BG_DARKEST, highlightbackground=Theme.BORDER_DARK, highlightthickness=1)
        right_split.add(editor_frame, minsize=240, stretch="always")

        self.notes_text = tk.Text(
            editor_frame,
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MAIN,
            insertbackground=Theme.RED_PRIMARY,
            selectbackground=Theme.RED_MUTED,
            selectforeground=Theme.TEXT_WHITE,
            font=("Segoe UI", 10),
            wrap=tk.WORD,
            bd=0,
            padx=10,
            pady=10
        )
        sb_notes_y = ttk.Scrollbar(editor_frame, orient=tk.VERTICAL, command=self.notes_text.yview)
        self.notes_text.configure(yscrollcommand=sb_notes_y.set)

        sb_notes_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure syntax highlight tags for Markdown
        self.notes_text.tag_config("h1", foreground=Theme.RED_PRIMARY, font=("Segoe UI", 13, "bold"))
        self.notes_text.tag_config("h2", foreground=Theme.RED_GLOW, font=("Segoe UI", 11, "bold"))
        self.notes_text.tag_config("h3", foreground="#facc15", font=("Segoe UI", 10, "bold"))
        self.notes_text.tag_config("keyword", foreground="#38bdf8", font=("Segoe UI", 10, "bold"))
        self.notes_text.tag_config("bullet", foreground=Theme.TEXT_WHITE)
        self.notes_text.tag_config("quote", foreground=Theme.TEXT_MUTED, font=("Segoe UI", 9, "italic"))

        # Lower: Execution Console Log
        self.analyze_console = ConsoleLogWidget(right_split, title="AI Template Analyzer Execution & Reasoning Log", height=8)
        right_split.add(self.analyze_console, minsize=140, stretch="always")

        # Initial data load
        self._refresh_analyze_templates_list()

    def _refresh_analyze_templates_list(self):
        """Scans data/*.pptx templates and updates Treeview and NOTE.md editor."""
        for item in self.tree_analyze_templates.get_children():
            self.tree_analyze_templates.delete(item)

        pptx_files = sorted(list(DATA_DIR.glob("*.pptx")))
        templates = [f for f in pptx_files if not f.name.endswith("_generated.pptx")]
        analyzed_map = get_analyzed_templates()

        analyzed_count = 0
        first_item_id = None

        for tpl in templates:
            try:
                prs = Presentation(str(tpl))
                slide_count = len(prs.slides)
            except Exception:
                slide_count = "?"

            is_analyzed = tpl.name in analyzed_map
            if is_analyzed:
                analyzed_count += 1
                status_str = "✓ Analyzed"
                tag = "analyzed"
                style_str = analyzed_map[tpl.name].get("style", "Custom")[:22]
                purpose_str = analyzed_map[tpl.name].get("purpose", "Ready")[:30]
            else:
                status_str = "○ Pending"
                tag = "pending"
                style_str = "Not analyzed"
                purpose_str = "Run AI Analyzer"

            item_id = self.tree_analyze_templates.insert(
                "",
                tk.END,
                text=tpl.name,
                values=(f"{slide_count} slides", status_str, style_str, purpose_str),
                tags=(tag,)
            )
            if not first_item_id:
                first_item_id = item_id

        # Reload NOTE.md into editor
        self._reload_analyze_notes()
        self.notes_status_var.set(f"data/NOTE.md • {analyzed_count}/{len(templates)} Templates Analyzed")
        self.analyze_progress_status_var.set(f"Ready • {analyzed_count}/{len(templates)} templates documented in NOTE.md")

        if first_item_id and not self.analyze_selected_file_path:
            self.tree_analyze_templates.selection_set(first_item_id)
            self._on_analyze_tree_selected()

    def _on_analyze_tree_selected(self, event=None):
        """Handles template selection in treeview and updates metadata & previews."""
        sel = self.tree_analyze_templates.selection()
        if not sel:
            return

        tpl_name = self.tree_analyze_templates.item(sel[0], "text")
        tpl_path = DATA_DIR / tpl_name
        if not tpl_path.exists():
            return

        self.analyze_selected_file_path = str(tpl_path)

        # Update metadata card
        analyzed_map = get_analyzed_templates()
        info = analyzed_map.get(tpl_name)

        try:
            prs = Presentation(str(tpl_path))
            slide_w_in = round(prs.slide_width / 914400, 2)
            slide_h_in = round(prs.slide_height / 914400, 2)
            dim_str = f"{slide_w_in}\" x {slide_h_in}\""
            total_slides = len(prs.slides)
        except Exception:
            dim_str = "Unknown"
            total_slides = "?"

        self.analyze_meta_text.config(state="normal")
        self.analyze_meta_text.delete("1.0", tk.END)
        self.analyze_meta_text.insert(tk.END, f"Template: {tpl_name}\n", "header")
        self.analyze_meta_text.insert(tk.END, f"Slides: {total_slides}  |  Dimensions: {dim_str}\n\n")

        if info:
            self.analyze_meta_text.insert(tk.END, f"🎯 Purpose: {info.get('purpose')}\n")
            self.analyze_meta_text.insert(tk.END, f"🎨 Style / Feel: {info.get('style')}\n")
            self.analyze_meta_text.insert(tk.END, f"📝 Brief: {info.get('brief')}\n")
        else:
            self.analyze_meta_text.insert(tk.END, "Status: ○ Not yet analyzed by AI Agent.\nClick '⚡ Analyze Selected Template' to generate design notes.")

        self.analyze_meta_text.config(state="disabled")

        # Load slide previews
        self._load_analyze_previews_async(str(tpl_path))

    def _load_analyze_previews_async(self, pptx_path: str):
        """Loads slide preview images for selected template asynchronously."""
        def worker():
            try:
                res = render_pptx_file_previews(pptx_path, target_width_px=500, return_engine_info=True)
                if isinstance(res, tuple):
                    imgs, engine_name = res
                else:
                    imgs, engine_name = res, "Renderer"
                self.after(0, lambda: self._apply_analyze_previews(imgs, engine_name))
            except Exception as e:
                self.after(0, lambda: self.analyze_console.log(f"Preview render notice: {e}", "dim"))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_analyze_previews(self, images: List[Image.Image], engine_name: str):
        self.analyze_preview_pil_images = images
        self.analyze_preview_engine_name = engine_name
        self.analyze_current_slide_idx = 0
        self._update_analyze_preview_display()

    def _on_analyze_preview_resize(self):
        if self.analyze_preview_pil_images and self.analyze_current_slide_idx < len(self.analyze_preview_pil_images):
            self._update_analyze_preview_display()

    def _update_analyze_preview_display(self):
        if not self.analyze_preview_pil_images:
            self.analyze_preview_label.config(image="", text="No preview available.")
            self.analyze_slide_counter_var.set("0 / 0")
            self.btn_analyze_prev.set_state("disabled")
            self.btn_analyze_next.set_state("disabled")
            return

        total = len(self.analyze_preview_pil_images)
        self.analyze_slide_counter_var.set(f"{self.analyze_current_slide_idx + 1} / {total}")
        self.btn_analyze_prev.set_state("normal" if self.analyze_current_slide_idx > 0 else "disabled")
        self.btn_analyze_next.set_state("normal" if self.analyze_current_slide_idx < total - 1 else "disabled")

        box_w = max(80, self.analyze_preview_display_box.winfo_width() - 12)
        box_h = max(80, self.analyze_preview_display_box.winfo_height() - 12)

        raw_img = self.analyze_preview_pil_images[self.analyze_current_slide_idx]
        img_w, img_h = raw_img.size

        scale = min(box_w / max(img_w, 1), box_h / max(img_h, 1))
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))

        try:
            resized_img = raw_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self._analyze_curr_tk_img = ImageTk.PhotoImage(resized_img)
            self.analyze_preview_label.config(image=self._analyze_curr_tk_img, text="")
        except Exception:
            pass

    def _analyze_prev_slide(self):
        if self.analyze_current_slide_idx > 0:
            self.analyze_current_slide_idx -= 1
            self._update_analyze_preview_display()

    def _analyze_next_slide(self):
        if self.analyze_preview_pil_images and self.analyze_current_slide_idx < len(self.analyze_preview_pil_images) - 1:
            self.analyze_current_slide_idx += 1
            self._update_analyze_preview_display()

    def _analyze_open_selected_in_ppt(self):
        if self.analyze_selected_file_path and Path(self.analyze_selected_file_path).exists():
            self._launch_file(self.analyze_selected_file_path)
        else:
            messagebox.showwarning("Warning", "No template file selected.")

    def _analyze_selected_template(self):
        """Runs AI template analysis on the currently selected template."""
        if not self.analyze_selected_file_path:
            messagebox.showwarning("Warning", "Please select a template from the list first.")
            return

        tpl_path = Path(self.analyze_selected_file_path)
        self.btn_analyze_sel.set_state("disabled")
        self.btn_analyze_all.set_state("disabled")
        self.analyze_progress_status_var.set(f"Analyzing {tpl_path.name} with 9Router AI Agent...")
        self.analyze_progress_var.set(25.0)

        def log_cb(msg: str):
            self.after(0, lambda: self.analyze_console.log(msg))

        def worker():
            try:
                log_cb(f"\n=======================================================")
                log_cb(f"[*] Starting AI Template Analysis: {tpl_path.name}")
                log_cb(f"[*] Using Model: {Config.NINEROUTER_CHAT_MODEL}")
                log_cb(f"=======================================================")
                
                result = analyze_template(tpl_path, log_cb=log_cb, save_to_file=True)
                
                log_cb(f"\n[✓] Analysis complete for {tpl_path.name}.")
                log_cb(f"    Purpose: {result.get('purpose')}")
                log_cb(f"    Style: {result.get('style')}")
                log_cb(f"[✓] Saved intelligence note to {NOTE_FILE}.")

                self.after(0, lambda: self.analyze_progress_var.set(100.0))
                self.after(0, lambda: self.analyze_progress_status_var.set(f"Analyzed {tpl_path.name} successfully."))
                self.after(0, self._refresh_analyze_templates_list)
            except Exception as e:
                log_cb(f"\n[!] Error during template analysis: {e}")
                self.after(0, lambda: messagebox.showerror("Analysis Error", str(e)))
            finally:
                self.after(0, lambda: self.btn_analyze_sel.set_state("normal"))
                self.after(0, lambda: self.btn_analyze_all.set_state("normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _analyze_all_templates_batch(self):
        """Sequentially analyzes all templates in data/ and updates NOTE.md."""
        pptx_files = sorted(list(DATA_DIR.glob("*.pptx")))
        templates = [f for f in pptx_files if not f.name.endswith("_generated.pptx")]

        if not templates:
            messagebox.showwarning("Warning", "No PPTX templates found in data folder.")
            return

        if not messagebox.askyesno(
            "Batch Template Analysis",
            f"Analyze all {len(templates)} templates with 9Router AI Agent?\nThis will generate comprehensive design notes in data/NOTE.md."
        ):
            return

        self.btn_analyze_sel.set_state("disabled")
        self.btn_analyze_all.set_state("disabled")
        self.analyze_progress_var.set(0.0)

        def log_cb(msg: str):
            self.after(0, lambda: self.analyze_console.log(msg))

        def progress_cb(current: int, total: int, current_name: str):
            pct = (current / total) * 100.0
            self.after(0, lambda: self.analyze_progress_var.set(pct))
            self.after(0, lambda: self.analyze_progress_status_var.set(f"Analyzing [{current}/{total}]: {current_name}"))

        def worker():
            try:
                log_cb(f"\n=======================================================")
                log_cb(f"[*] Starting Batch Template Analysis ({len(templates)} templates)")
                log_cb(f"[*] Output destination: {NOTE_FILE}")
                log_cb(f"=======================================================")

                analyze_all_templates(DATA_DIR, progress_cb=progress_cb, log_cb=log_cb)

                self.after(0, lambda: self.analyze_progress_var.set(100.0))
                self.after(0, lambda: self.analyze_progress_status_var.set(f"Completed analysis of all {len(templates)} templates."))
                self.after(0, self._refresh_analyze_templates_list)
                self.after(0, lambda: messagebox.showinfo("Success", f"All {len(templates)} templates successfully analyzed!\nNotes saved to: {NOTE_FILE}"))
            except Exception as e:
                log_cb(f"\n[!] Batch analysis error: {e}")
                self.after(0, lambda: messagebox.showerror("Batch Error", str(e)))
            finally:
                self.after(0, lambda: self.btn_analyze_sel.set_state("normal"))
                self.after(0, lambda: self.btn_analyze_all.set_state("normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _save_analyze_notes(self):
        """Saves manual edits made to NOTE.md in the text editor."""
        content = self.notes_text.get("1.0", tk.END).strip()
        save_notes(content, NOTE_FILE)
        self._highlight_notes_syntax()
        self.analyze_console.log(f"[✓] Manually saved changes to {NOTE_FILE}.", "success")
        messagebox.showinfo("Saved", f"Template Intelligence Notes successfully saved to:\n{NOTE_FILE}")
        self._refresh_analyze_templates_list()

    def _reload_analyze_notes(self):
        """Reloads NOTE.md content from disk into the text editor."""
        content = load_notes(NOTE_FILE)
        self.notes_text.delete("1.0", tk.END)
        if content:
            self.notes_text.insert(tk.END, content)
            self._highlight_notes_syntax()
        else:
            self.notes_text.insert(
                tk.END,
                "# PPTX Jahat — Template Intelligence & Design Notes\n\n"
                "No template notes found yet.\n\n"
                "Click '⚡ Analyze Selected Template' or '🚀 Analyze All Templates (Batch)' above to generate AI notes for your PPTX templates!"
            )

    def _copy_analyze_notes(self):
        """Copies editor content to clipboard."""
        content = self.notes_text.get("1.0", tk.END).strip()
        self.clipboard_clear()
        self.clipboard_append(content)
        self.analyze_console.log("[✓] Notes copied to clipboard.", "info")

    def _highlight_notes_syntax(self):
        """Applies syntax highlighting tags to markdown content."""
        # Remove existing tags
        for tag in ["h1", "h2", "h3", "keyword", "quote"]:
            self.notes_text.tag_remove(tag, "1.0", tk.END)

        lines = self.notes_text.get("1.0", tk.END).split("\n")
        for line_idx, line in enumerate(lines, start=1):
            stripped = line.strip()
            if stripped.startswith("# "):
                self.notes_text.tag_add("h1", f"{line_idx}.0", f"{line_idx}.end")
            elif stripped.startswith("## "):
                self.notes_text.tag_add("h2", f"{line_idx}.0", f"{line_idx}.end")
            elif stripped.startswith("### "):
                self.notes_text.tag_add("h3", f"{line_idx}.0", f"{line_idx}.end")
            elif stripped.startswith("> "):
                self.notes_text.tag_add("quote", f"{line_idx}.0", f"{line_idx}.end")
            
            # Highlight key section keywords
            for kw in ["🎯 Purpose", "🎨 Style", "💡 Core Concept", "📝 Content Brief", "📊 Slide Inventory", "🤖 AI Selection", "When to Choose"]:
                if kw in line:
                    col_start = line.find(kw)
                    col_end = col_start + len(kw)
                    self.notes_text.tag_add("keyword", f"{line_idx}.{col_start}", f"{line_idx}.{col_end}")

    # -------------------------------------------------------------
    # TAB 3: Deck & Template Manager (Manage generated PPTX & Reference Templates)
    # -------------------------------------------------------------
    def _setup_manager_tab(self):
        paned = tk.PanedWindow(
            self.tab_manager,
            orient=tk.HORIZONTAL,
            bg=Theme.BG_DARKEST,
            bd=0,
            sashwidth=4,
            sashrelief="flat"
        )
        paned.pack(fill=tk.BOTH, expand=True)

        # --- LEFT PANEL: Dual File Treeview for Output Presentations & Reference Templates ---
        left_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(left_container, minsize=520, stretch="always")

        # Top Control Bar for Importing and Refreshing
        top_ctrl_card = ModernCard(
            left_container,
            title="PRESENTATION & TEMPLATE MANAGER",
            subtitle="Manage generated decks (data/output) and reference styles (data/*.pptx)",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        top_ctrl_card.pack(fill=tk.X, pady=(0, 10))

        btn_bar = tk.Frame(top_ctrl_card.body, bg=Theme.BG_SURFACE)
        btn_bar.pack(fill=tk.X, pady=2)

        StyledActionBtn(
            btn_bar,
            text="📥 Import Reference PPTX",
            command=self._mgr_import_template,
            is_primary=True,
            padx=10,
            pady=5
        ).pack(side=tk.LEFT, padx=(0, 6))

        StyledActionBtn(
            btn_bar,
            text="📂 Open Output Folder",
            command=lambda: self._reveal_in_explorer(str(OUTPUT_DIR)),
            is_primary=False,
            padx=10,
            pady=5
        ).pack(side=tk.LEFT, padx=4)

        StyledActionBtn(
            btn_bar,
            text="📁 Open Templates Folder",
            command=lambda: self._reveal_in_explorer(str(DATA_DIR)),
            is_primary=False,
            padx=10,
            pady=5
        ).pack(side=tk.LEFT, padx=4)

        StyledActionBtn(
            btn_bar,
            text="🔄 Refresh All",
            command=self._refresh_manager_lists,
            is_primary=False,
            padx=10,
            pady=5
        ).pack(side=tk.RIGHT)

        # Split Treeviews (Generated Decks vs Reference Templates)
        lists_frame = tk.Frame(left_container, bg=Theme.BG_MAIN)
        lists_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Generated Presentations Sub-frame
        gen_frame = ModernCard(
            lists_frame,
            title="GENERATED PRESENTATIONS (data/output)",
            subtitle="Recently synthesized slide decks",
            accent_color=Theme.RED_PRIMARY
        )
        gen_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        tree_gen_container = tk.Frame(gen_frame.body, bg=Theme.BG_DARKEST, highlightbackground=Theme.BORDER_DARK, highlightthickness=1)
        tree_gen_container.pack(fill=tk.BOTH, expand=True)

        self.tree_generated = ttk.Treeview(
            tree_gen_container,
            columns=("name", "size", "modified", "path"),
            show="headings",
            selectmode="browse",
            height=6
        )
        self.tree_generated.heading("name", text="Presentation Name")
        self.tree_generated.heading("size", text="File Size")
        self.tree_generated.heading("modified", text="Date Modified")
        self.tree_generated.heading("path", text="Full Path")

        self.tree_generated.column("name", width=220, anchor="w")
        self.tree_generated.column("size", width=80, anchor="center")
        self.tree_generated.column("modified", width=140, anchor="center")
        self.tree_generated.column("path", width=0, stretch=False) # Hidden full path

        scroll_gen_y = ttk.Scrollbar(tree_gen_container, orient=tk.VERTICAL, command=self.tree_generated.yview)
        self.tree_generated.configure(yscrollcommand=scroll_gen_y.set)
        scroll_gen_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_generated.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_generated.bind("<<TreeviewSelect>>", lambda e: self._on_tree_item_selected(self.tree_generated, "generated"))
        self.tree_generated.bind("<Double-1>", lambda e: self._mgr_open_selected_in_ppt())

        # 2. Reference Templates Sub-frame
        ref_frame = ModernCard(
            lists_frame,
            title="REFERENCE TEMPLATES (data/*.pptx)",
            subtitle="Source templates for AI component infilling & layout matching",
            accent_color=Theme.BORDER_LIGHT
        )
        ref_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        tree_ref_container = tk.Frame(ref_frame.body, bg=Theme.BG_DARKEST, highlightbackground=Theme.BORDER_DARK, highlightthickness=1)
        tree_ref_container.pack(fill=tk.BOTH, expand=True)

        self.tree_reference = ttk.Treeview(
            tree_ref_container,
            columns=("name", "size", "modified", "path"),
            show="headings",
            selectmode="browse",
            height=6
        )
        self.tree_reference.heading("name", text="Template File")
        self.tree_reference.heading("size", text="File Size")
        self.tree_reference.heading("modified", text="Date Modified")
        self.tree_reference.heading("path", text="Full Path")

        self.tree_reference.column("name", width=220, anchor="w")
        self.tree_reference.column("size", width=80, anchor="center")
        self.tree_reference.column("modified", width=140, anchor="center")
        self.tree_reference.column("path", width=0, stretch=False) # Hidden full path

        scroll_ref_y = ttk.Scrollbar(tree_ref_container, orient=tk.VERTICAL, command=self.tree_reference.yview)
        self.tree_reference.configure(yscrollcommand=scroll_ref_y.set)
        scroll_ref_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_reference.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_reference.bind("<<TreeviewSelect>>", lambda e: self._on_tree_item_selected(self.tree_reference, "reference"))
        self.tree_reference.bind("<Double-1>", lambda e: self._mgr_open_selected_in_ppt())

        # --- RIGHT PANEL: File Details, Quick Actions & Live Visual Inspector ---
        right_container = tk.Frame(paned, bg=Theme.BG_MAIN, padx=5, pady=5)
        paned.add(right_container, minsize=480, stretch="always")

        details_card = ModernCard(
            right_container,
            title="DECK INSPECTOR & ACTIONS",
            subtitle="Examine slides, rename, duplicate, or launch presentations",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        details_card.pack(fill=tk.BOTH, expand=True)

        # Selected File Meta Info Box
        info_box = tk.Frame(details_card.body, bg=Theme.BG_SURFACE, padx=10, pady=8)
        info_box.pack(fill=tk.X, side=tk.TOP, pady=(0, 6))

        self.mgr_file_name_var = tk.StringVar(value="Select a presentation or template from the left list")
        lbl_file_name = tk.Label(
            info_box,
            textvariable=self.mgr_file_name_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_WHITE,
            font=Theme.FONT_TITLE,
            anchor="w"
        )
        lbl_file_name.pack(anchor="w")

        self.mgr_file_details_var = tk.StringVar(value="Path: None | Size: -- | Type: --")
        lbl_file_details = tk.Label(
            info_box,
            textvariable=self.mgr_file_details_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_CAPTION,
            anchor="w"
        )
        lbl_file_details.pack(anchor="w", pady=(2, 0))

        # Action Buttons Toolbar for Selected Item
        action_toolbar = tk.Frame(details_card.body, bg=Theme.BG_SURFACE, padx=6, pady=6)
        action_toolbar.pack(fill=tk.X, side=tk.TOP, pady=(0, 8))

        self.mgr_btn_open_ppt = StyledActionBtn(
            action_toolbar,
            text="📊 Open in PowerPoint",
            command=self._mgr_open_selected_in_ppt,
            is_primary=True,
            padx=10,
            pady=4
        )
        self.mgr_btn_open_ppt.set_state("disabled")
        self.mgr_btn_open_ppt.pack(side=tk.LEFT, padx=3)

        self.mgr_btn_verify = StyledActionBtn(
            action_toolbar,
            text="🛡️ Verify & Fix",
            command=self._mgr_verify_and_fix_selected,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.mgr_btn_verify.set_state("disabled")
        self.mgr_btn_verify.pack(side=tk.LEFT, padx=3)

        self.mgr_btn_reveal = StyledActionBtn(
            action_toolbar,
            text="📁 Reveal File",
            command=self._mgr_reveal_selected,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.mgr_btn_reveal.set_state("disabled")
        self.mgr_btn_reveal.pack(side=tk.LEFT, padx=3)

        self.mgr_btn_duplicate = StyledActionBtn(
            action_toolbar,
            text="📋 Duplicate",
            command=self._mgr_duplicate_selected,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.mgr_btn_duplicate.set_state("disabled")
        self.mgr_btn_duplicate.pack(side=tk.LEFT, padx=3)

        self.mgr_btn_rename = StyledActionBtn(
            action_toolbar,
            text="✏️ Rename",
            command=self._mgr_rename_selected,
            is_primary=False,
            padx=10,
            pady=4
        )
        self.mgr_btn_rename.set_state("disabled")
        self.mgr_btn_rename.pack(side=tk.LEFT, padx=3)

        self.mgr_btn_delete = StyledActionBtn(
            action_toolbar,
            text="🗑️ Delete",
            command=self._mgr_delete_selected,
            is_primary=False,
            is_danger=True,
            padx=10,
            pady=4
        )
        self.mgr_btn_delete.set_state("disabled")
        self.mgr_btn_delete.pack(side=tk.RIGHT, padx=3)

        # Slide Navigation Bar
        mgr_nav_toolbar = tk.Frame(details_card.body, bg=Theme.BG_SURFACE, pady=4)
        mgr_nav_toolbar.pack(fill=tk.X, side=tk.TOP)

        self.mgr_btn_prev_slide = StyledActionBtn(
            mgr_nav_toolbar,
            text="◀ Prev",
            command=self._mgr_prev_slide,
            is_primary=False,
            padx=8,
            pady=3
        )
        self.mgr_btn_prev_slide.set_state("disabled")
        self.mgr_btn_prev_slide.pack(side=tk.LEFT, padx=4)

        self.mgr_slide_counter_var = tk.StringVar(value="No preview loaded")
        self.mgr_preview_engine_badge_var = tk.StringVar(value="")

        mgr_center_info_box = tk.Frame(mgr_nav_toolbar, bg=Theme.BG_SURFACE)
        mgr_center_info_box.pack(side=tk.LEFT, expand=True)

        lbl_mgr_counter = tk.Label(
            mgr_center_info_box,
            textvariable=self.mgr_slide_counter_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_mgr_counter.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_mgr_preview_engine = tk.Label(
            mgr_center_info_box,
            textvariable=self.mgr_preview_engine_badge_var,
            bg=Theme.BADGE_BG_RED,
            fg=Theme.BADGE_TEXT_RED,
            font=Theme.FONT_CAPTION,
            padx=6,
            pady=1,
            highlightbackground=Theme.BADGE_BORDER_RED,
            highlightthickness=1
        )
        self.lbl_mgr_preview_engine.pack(side=tk.LEFT)

        self.mgr_btn_next_slide = StyledActionBtn(
            mgr_nav_toolbar,
            text="Next ▶",
            command=self._mgr_next_slide,
            is_primary=False,
            padx=8,
            pady=3
        )
        self.mgr_btn_next_slide.set_state("disabled")
        self.mgr_btn_next_slide.pack(side=tk.RIGHT, padx=4)

        # Slide Display Canvas Box
        self.mgr_preview_display_box = tk.Frame(
            details_card.body,
            bg=Theme.BG_DARKEST,
            highlightbackground=Theme.BORDER_DARK,
            highlightthickness=1
        )
        self.mgr_preview_display_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.mgr_preview_display_box.bind("<Configure>", lambda e: self._on_mgr_preview_resize())

        self.mgr_preview_label = tk.Label(
            self.mgr_preview_display_box,
            text="Select a presentation or template to inspect slide layouts.",
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MUTED,
            font=Theme.FONT_BODY,
            anchor="center"
        )
        self.mgr_preview_label.pack(fill=tk.BOTH, expand=True)

        self._refresh_manager_lists()

    def _refresh_manager_lists(self):
        """Reloads both generated presentations and template lists."""
        # 1. Clear treeviews
        for row in self.tree_generated.get_children():
            self.tree_generated.delete(row)
        for row in self.tree_reference.get_children():
            self.tree_reference.delete(row)

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)

        # 2. Populate Generated PPTX
        gen_files = sorted(OUTPUT_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
        for p in gen_files:
            try:
                st = p.stat()
                size_str = f"{st.st_size / 1024:.1f} KB" if st.st_size < 1024*1024 else f"{st.st_size / (1024*1024):.2f} MB"
                mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
                self.tree_generated.insert("", tk.END, values=(p.name, size_str, mtime_str, str(p.resolve())))
            except Exception:
                pass

        # 3. Populate Reference Templates (exclude output folder pptxs)
        ref_files = sorted([p for p in DATA_DIR.glob("*.pptx") if not p.name.endswith("_generated.pptx") and p.parent.resolve() == DATA_DIR.resolve()], key=lambda p: p.name)
        for p in ref_files:
            try:
                st = p.stat()
                size_str = f"{st.st_size / 1024:.1f} KB" if st.st_size < 1024*1024 else f"{st.st_size / (1024*1024):.2f} MB"
                mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
                self.tree_reference.insert("", tk.END, values=(p.name, size_str, mtime_str, str(p.resolve())))
            except Exception:
                pass

        self._refresh_templates()

    def _on_tree_item_selected(self, tree: ttk.Treeview, source_type: str):
        sel = tree.selection()
        if not sel:
            return

        # Deselect the other tree
        other_tree = self.tree_reference if tree == self.tree_generated else self.tree_generated
        for s in other_tree.selection():
            other_tree.selection_remove(s)

        item = tree.item(sel[0])
        values = item["values"]
        if not values or len(values) < 4:
            return

        file_name, size_str, modified_str, file_path = values[0], values[1], values[2], values[3]
        self.mgr_selected_file_path = file_path

        # Update labels & enable buttons
        self.mgr_file_name_var.set(f"📄 {file_name}")
        type_label = "Generated Deck" if source_type == "generated" else "Reference Template"
        self.mgr_file_details_var.set(f"Type: {type_label} | Size: {size_str} | Modified: {modified_str}\nPath: {file_path}")

        self.mgr_btn_open_ppt.set_state("normal")
        self.mgr_btn_verify.set_state("normal")
        self.mgr_btn_reveal.set_state("normal")
        self.mgr_btn_duplicate.set_state("normal")
        self.mgr_btn_rename.set_state("normal")
        self.mgr_btn_delete.set_state("normal")

        # Load previews asynchronously in background thread
        self._load_mgr_previews_async(file_path)

    def _load_mgr_previews_async(self, file_path: str):
        self.mgr_slide_counter_var.set("Loading previews...")
        self.mgr_preview_engine_badge_var.set("")
        self.lbl_mgr_preview_engine.config(bg=Theme.BG_SURFACE, highlightthickness=0)
        self.mgr_preview_label.config(image="", text="Rendering slide previews...")

        def worker():
            try:
                res = render_pptx_file_previews(file_path, target_width_px=750, return_engine_info=True)
                if isinstance(res, tuple):
                    imgs, engine_name = res
                else:
                    imgs, engine_name = res, "Renderer"
                self.after(50, lambda: self._apply_mgr_previews(imgs, engine_name))
            except Exception as ex:
                err_msg = str(ex)
                self.after(50, lambda: self.mgr_slide_counter_var.set("Preview Error"))
                self.after(50, lambda msg=err_msg: self.mgr_preview_label.config(text=f"Could not render preview: {msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_mgr_previews(self, imgs: List[Image.Image], engine_name: str = ""):
        self.mgr_preview_pil_images = imgs
        self.mgr_preview_engine_name = engine_name
        self.mgr_current_slide_idx = 0
        self._update_mgr_preview_display()

    def _on_mgr_preview_resize(self):
        if self.mgr_preview_pil_images and self.mgr_current_slide_idx < len(self.mgr_preview_pil_images):
            self._update_mgr_preview_display()

    def _update_mgr_preview_display(self):
        if not self.mgr_preview_pil_images:
            self.mgr_preview_label.config(image="", text="No slide preview available.")
            self.mgr_slide_counter_var.set("0 / 0")
            self.mgr_preview_engine_badge_var.set("")
            self.lbl_mgr_preview_engine.config(bg=Theme.BG_SURFACE, highlightthickness=0)
            self.mgr_btn_prev_slide.set_state("disabled")
            self.mgr_btn_next_slide.set_state("disabled")
            return

        total = len(self.mgr_preview_pil_images)
        self.mgr_slide_counter_var.set(f"Slide {self.mgr_current_slide_idx + 1} of {total}")

        # Engine badge indicator
        if self.mgr_preview_engine_name:
            if "PowerPoint" in self.mgr_preview_engine_name:
                self.mgr_preview_engine_badge_var.set("⚡ Native PowerPoint")
                self.lbl_mgr_preview_engine.config(bg="#1e3a29", fg="#4ade80", highlightbackground="#22c55e", highlightthickness=1)
            else:
                self.mgr_preview_engine_badge_var.set("🎨 Pure PIL Engine")
                self.lbl_mgr_preview_engine.config(bg=Theme.BADGE_BG_RED, fg=Theme.BADGE_TEXT_RED, highlightbackground=Theme.BADGE_BORDER_RED, highlightthickness=1)
        else:
            self.mgr_preview_engine_badge_var.set("")
            self.lbl_mgr_preview_engine.config(bg=Theme.BG_SURFACE, highlightthickness=0)

        box_w = max(100, self.mgr_preview_display_box.winfo_width() - 20)
        box_h = max(100, self.mgr_preview_display_box.winfo_height() - 20)

        raw_img = self.mgr_preview_pil_images[self.mgr_current_slide_idx]
        img_w, img_h = raw_img.size

        scale = min(box_w / img_w, box_h / img_h, 1.0)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))

        resized = raw_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        if resized.mode == "RGBA":
            bg = Image.new("RGBA", resized.size, (255, 255, 255, 255))
            bg.alpha_composite(resized)
            resized = bg.convert("RGB")
        self._mgr_current_tk_img = ImageTk.PhotoImage(resized)

        self.mgr_preview_label.config(image=self._mgr_current_tk_img, text="")

        self.mgr_btn_prev_slide.set_state("normal" if self.mgr_current_slide_idx > 0 else "disabled")
        self.mgr_btn_next_slide.set_state("normal" if self.mgr_current_slide_idx < total - 1 else "disabled")

    def _mgr_prev_slide(self):
        if self.mgr_current_slide_idx > 0:
            self.mgr_current_slide_idx -= 1
            self._update_mgr_preview_display()

    def _mgr_next_slide(self):
        if self.mgr_current_slide_idx < len(self.mgr_preview_pil_images) - 1:
            self.mgr_current_slide_idx += 1
            self._update_mgr_preview_display()

    def _mgr_open_selected_in_ppt(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            messagebox.showwarning("Warning", "No file selected.")
            return
        self._launch_file(self.mgr_selected_file_path)

    def _mgr_verify_and_fix_selected(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            messagebox.showwarning("Warning", "No file selected.")
            return
        target_path = self.mgr_selected_file_path
        self.badge_status.set_text("● VERIFYING...", fg_color=Theme.TEXT_WHITE, bg_color=Theme.RED_PRIMARY)
        self.status_left_var.set(f"Verifying integrity for {Path(target_path).name}...")

        def worker():
            is_ok, final_p = verify_and_auto_heal_pptx(target_path)
            self.after(50, lambda: self._on_mgr_verify_done(is_ok, final_p))

        threading.Thread(target=worker, daemon=True).start()

    def _on_mgr_verify_done(self, is_ok: bool, final_p: str):
        self.badge_status.set_text("● SYSTEM READY", fg_color=Theme.BADGE_TEXT_RED, bg_color=Theme.BADGE_BG_RED)
        self.status_left_var.set(f"Verification complete: {Path(final_p).name}")
        self._load_mgr_previews_async(final_p)
        if is_ok:
            messagebox.showinfo("Verification Passed", f"PPTX presentation is valid and clean!\n{final_p}")
        else:
            messagebox.showwarning("Notice", f"Applied auto-repairs to presentation.\n{final_p}")

    def _mgr_reveal_selected(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            messagebox.showwarning("Warning", "No file selected.")
            return
        self._reveal_in_explorer(self.mgr_selected_file_path)

    def _mgr_duplicate_selected(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            return
        orig_p = Path(self.mgr_selected_file_path)
        copy_p = orig_p.parent / f"{orig_p.stem}_copy{orig_p.suffix}"
        idx = 1
        while copy_p.exists():
            copy_p = orig_p.parent / f"{orig_p.stem}_copy{idx}{orig_p.suffix}"
            idx += 1

        try:
            shutil.copy2(orig_p, copy_p)
            self._refresh_manager_lists()
            messagebox.showinfo("Duplicated", f"Created copy at:\n{copy_p.name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to duplicate file: {str(e)}")

    def _mgr_rename_selected(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            return
        orig_p = Path(self.mgr_selected_file_path)

        # Dialog for new name
        win = tk.Toplevel(self)
        win.title("Rename Presentation")
        win.geometry("450x160")
        win.configure(bg=Theme.BG_DARKEST)
        win.transient(self)
        win.grab_set()

        lbl = tk.Label(win, text=f"Rename '{orig_p.name}':", bg=Theme.BG_DARKEST, fg=Theme.TEXT_WHITE, font=Theme.FONT_TITLE)
        lbl.pack(anchor="w", padx=20, pady=(15, 6))

        var = tk.StringVar(value=orig_p.name)
        entry = ttk.Entry(win, textvariable=var, font=("Segoe UI", 10))
        entry.pack(fill=tk.X, padx=20, pady=4)
        entry.focus_set()
        entry.selection_range(0, len(orig_p.stem))

        btn_box = tk.Frame(win, bg=Theme.BG_DARKEST)
        btn_box.pack(fill=tk.X, padx=20, pady=15)

        def do_rename():
            new_name = var.get().strip()
            if not new_name:
                return
            if not new_name.endswith(".pptx"):
                new_name += ".pptx"
            target_p = orig_p.parent / new_name
            if target_p.exists() and target_p != orig_p:
                messagebox.showerror("Error", "A file with this name already exists.", parent=win)
                return
            try:
                orig_p.rename(target_p)
                win.destroy()
                self._refresh_manager_lists()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to rename file: {str(e)}", parent=win)

        StyledActionBtn(btn_box, text="Save Name", command=do_rename, is_primary=True).pack(side=tk.LEFT, padx=(0, 6))
        StyledActionBtn(btn_box, text="Cancel", command=win.destroy, is_primary=False).pack(side=tk.LEFT)

    def _mgr_delete_selected(self):
        if not self.mgr_selected_file_path or not Path(self.mgr_selected_file_path).exists():
            return
        orig_p = Path(self.mgr_selected_file_path)

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to permanently delete:\n{orig_p.name}?"):
            try:
                orig_p.unlink()
                self.mgr_selected_file_path = None
                self.mgr_file_name_var.set("Select a presentation or template from the left list")
                self.mgr_file_details_var.set("Path: None | Size: -- | Type: --")
                self.mgr_preview_pil_images = []
                self.mgr_preview_label.config(image="", text="Presentation deleted.")
                self.mgr_slide_counter_var.set("0 / 0")

                self.mgr_btn_open_ppt.set_state("disabled")
                self.mgr_btn_reveal.set_state("disabled")
                self.mgr_btn_duplicate.set_state("disabled")
                self.mgr_btn_rename.set_state("disabled")
                self.mgr_btn_delete.set_state("disabled")

                self._refresh_manager_lists()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete file: {str(e)}")

    def _mgr_import_template(self):
        files = filedialog.askopenfilenames(filetypes=[("PowerPoint Presentation", "*.pptx")])
        if not files:
            return

        imported_count = 0
        for f in files:
            src = Path(f)
            dest = DATA_DIR / src.name
            try:
                shutil.copy2(src, dest)
                imported_count += 1
            except Exception as e:
                messagebox.showerror("Error", f"Could not import {src.name}: {str(e)}")

        self._refresh_manager_lists()
        messagebox.showinfo("Import Complete", f"Successfully imported {imported_count} template(s) into data/.")

    # -------------------------------------------------------------
    # TAB 3: Components Catalog & Inspection
    # -------------------------------------------------------------
    def _setup_components_tab(self):
        container = tk.Frame(self.tab_components, bg=Theme.BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True)

        card = ModernCard(
            container,
            title="TEMPLATE REPOSITORY & COMPONENT CATALOG",
            subtitle="Inspect extracted shapes, layout containers, and design primitives",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        card.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        top_bar = tk.Frame(card.body, bg=Theme.BG_SURFACE, pady=4)
        top_bar.pack(fill=tk.X, side=tk.TOP)

        btn_scan = StyledActionBtn(
            top_bar,
            text="🔍 Scan & Extract Templates (data/*.pptx)",
            command=self._run_extraction,
            is_primary=True,
            padx=12,
            pady=5
        )
        btn_scan.pack(side=tk.LEFT, padx=(0, 8))

        btn_reload = StyledActionBtn(
            top_bar,
            text="🔄 Reload Catalog JSON",
            command=self._load_components_json,
            is_primary=False,
            padx=12,
            pady=5
        )
        btn_reload.pack(side=tk.LEFT)

        self.comp_stats_var = tk.StringVar(value="Components: Loading...")
        lbl_stats = tk.Label(
            top_bar,
            textvariable=self.comp_stats_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_RED,
            font=Theme.FONT_TITLE
        )
        lbl_stats.pack(side=tk.RIGHT, padx=6)

        # Editor / Display Area
        text_frame = tk.Frame(card.body, bg=Theme.BG_DARKEST, highlightbackground=Theme.BORDER_DARK, highlightthickness=1)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.comp_text = tk.Text(
            text_frame,
            bg=Theme.BG_DARKEST,
            fg=Theme.TEXT_MAIN,
            insertbackground=Theme.RED_PRIMARY,
            selectbackground=Theme.RED_MUTED,
            selectforeground=Theme.TEXT_WHITE,
            font=Theme.FONT_LOG,
            wrap=tk.WORD,
            bd=0,
            padx=10,
            pady=10
        )
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.comp_text.yview)
        self.comp_text.configure(yscrollcommand=scroll.set)
        
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.comp_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._load_components_json()

    def _run_extraction(self):
        def worker():
            self.comp_text.config(state="normal")
            self.comp_text.delete("1.0", tk.END)
            self.comp_text.insert(tk.END, "[*] Scanning data/*.pptx templates and extracting components...\n")
            self.comp_text.config(state="disabled")
            
            catalog = extract_all_templates()
            count = len(catalog.get("all_components", []))
            
            self.comp_text.config(state="normal")
            self.comp_text.insert(tk.END, f"[✓] Extraction finished. Discovered {count} visual design components.\n\n")
            self.comp_text.insert(tk.END, json.dumps(catalog, indent=2))
            self.comp_text.config(state="disabled")
            self.comp_stats_var.set(f"Extracted Components: {count}")
            
        threading.Thread(target=worker, daemon=True).start()

    def _load_components_json(self):
        catalog = get_components_catalog()
        count = len(catalog.get("all_components", []))
        self.comp_stats_var.set(f"Extracted Components: {count}")
        
        self.comp_text.config(state="normal")
        self.comp_text.delete("1.0", tk.END)
        self.comp_text.insert(tk.END, json.dumps(catalog, indent=2))
        self.comp_text.config(state="disabled")

    # -------------------------------------------------------------
    # TAB 4: Autonomous AI Agent Terminal
    # -------------------------------------------------------------
    def _setup_agent_tab(self):
        container = tk.Frame(self.tab_agent, bg=Theme.BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True)

        card = ModernCard(
            container,
            title="AUTONOMOUS AGENT WORKSPACE",
            subtitle="Execute multi-step presentation queries, tool automation, and web research",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        card.pack(fill=tk.BOTH, expand=True)

        # Agent Log Terminal
        self.agent_console = ConsoleLogWidget(card.body, title="Agent Execution & Tool Calling Pipeline", height=18)
        self.agent_console.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Bottom Input Prompt Bar
        input_bar = tk.Frame(card.body, bg=Theme.BG_SURFACE, pady=6)
        input_bar.pack(fill=tk.X, side=tk.BOTTOM)

        tk.Label(input_bar, text="Instruction Prompt:", bg=Theme.BG_SURFACE, fg=Theme.TEXT_MAIN, font=Theme.FONT_BODY_BOLD).pack(side=tk.LEFT, padx=(0, 8))

        self.agent_prompt_var = tk.StringVar()
        self.agent_entry = ttk.Entry(input_bar, textvariable=self.agent_prompt_var, font=("Segoe UI", 10))
        self.agent_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.agent_entry.bind("<Return>", lambda e: self._send_agent_prompt())

        self.btn_send_agent = StyledActionBtn(
            input_bar,
            text="🚀 Execute Instruction",
            command=self._send_agent_prompt,
            is_primary=True,
            padx=14,
            pady=6
        )
        self.btn_send_agent.pack(side=tk.LEFT)

        self.agent = AIAgent()

    def _send_agent_prompt(self):
        prompt = self.agent_prompt_var.get().strip()
        if not prompt:
            return
        self.agent_prompt_var.set("")
        self.agent_console.log(f"\n[USER PROMPT]: {prompt}", "accent")

        self.btn_send_agent.set_state("disabled")
        self.badge_status.set_text("● AGENT RUNNING", fg_color=Theme.TEXT_WHITE, bg_color=Theme.RED_PRIMARY)
        self.status_left_var.set("AI Agent executing reasoning loop...")

        def worker():
            def log_cb(msg):
                self.agent_console.log(f"  {msg}")

            try:
                reply = self.agent.run(prompt, log_callback=log_cb)
                self.agent_console.log(f"\n[AGENT FINAL RESPONSE]:\n{reply}\n", "success")
                self.after(100, lambda: self.badge_status.set_text("● SYSTEM READY", fg_color=Theme.BADGE_TEXT_RED, bg_color=Theme.BADGE_BG_RED))
                self.after(100, lambda: self.status_left_var.set("Agent task completed."))
            except Exception as e:
                self.agent_console.log(f"Agent Error: {str(e)}", "error")
                self.after(100, lambda: self.badge_status.set_text("● AGENT ERROR", fg_color=Theme.TEXT_RED, bg_color=Theme.BADGE_BG_RED))
                self.after(100, lambda: self.status_left_var.set("Agent failed."))
            finally:
                self.after(100, lambda: self.btn_send_agent.set_state("normal"))

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------------------
    # TAB 5: Settings & Configuration
    # -------------------------------------------------------------
    def _setup_settings_tab(self):
        container = tk.Frame(self.tab_settings, bg=Theme.BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True)

        card = ModernCard(
            container,
            title="9ROUTER & API CONFIGURATION",
            subtitle="Configure environment endpoints, model routing, and API credentials",
            show_accent_stripe=True,
            accent_color=Theme.RED_PRIMARY
        )
        card.pack(fill=tk.BOTH, expand=True)

        self.env_vars = {}
        fields = [
            ("NINEROUTER_URL", Config.NINEROUTER_URL, False, "9Router Base Gateway URL (e.g. http://localhost:20128)"),
            ("NINEROUTER_KEY", Config.NINEROUTER_KEY, True, "API Key for 9Router or OpenAI-compatible backend"),
            ("NINEROUTER_CHAT_MODEL", Config.NINEROUTER_CHAT_MODEL, False, "Primary Chat/Reasoning Model (e.g. ag/gemini-3.7-flash-high, openai/gpt-4o)"),
            ("NINEROUTER_SEARCH_MODEL", Config.NINEROUTER_SEARCH_MODEL, False, "Search Provider/Model (e.g. tavily, exa, brave-search)"),
            ("NINEROUTER_FETCH_MODEL", Config.NINEROUTER_FETCH_MODEL, False, "Web Reader/Scraper Provider (e.g. jina-reader, firecrawl, exa)"),
            ("NINEROUTER_IMAGE_MODEL", Config.NINEROUTER_IMAGE_MODEL, False, "Slide Image Generation Model (e.g. gemini/gemini-3-pro-image-preview)"),
        ]

        form_frame = tk.Frame(card.body, bg=Theme.BG_SURFACE)
        form_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        for k, v, is_secret, desc in fields:
            row = tk.Frame(form_frame, bg=Theme.BG_SURFACE)
            row.pack(fill=tk.X, pady=8)

            lbl_box = tk.Frame(row, bg=Theme.BG_SURFACE, width=280)
            lbl_box.pack(side=tk.LEFT)
            tk.Label(lbl_box, text=k, bg=Theme.BG_SURFACE, fg=Theme.TEXT_WHITE, font=Theme.FONT_BODY_BOLD).pack(anchor="w")
            tk.Label(lbl_box, text=desc, bg=Theme.BG_SURFACE, fg=Theme.TEXT_MUTED, font=Theme.FONT_CAPTION).pack(anchor="w")

            var = tk.StringVar(value=v)
            self.env_vars[k] = var
            entry = ttk.Entry(row, textvariable=var, show="*" if is_secret else "")
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=12)

        btn_row = tk.Frame(card.body, bg=Theme.BG_SURFACE, pady=12)
        btn_row.pack(fill=tk.X, side=tk.BOTTOM)

        btn_save = StyledActionBtn(
            btn_row,
            text="💾 Save Configuration (.env) & Reload Engine",
            command=self._save_env_settings,
            is_primary=True
        )
        btn_save.pack(side=tk.LEFT)

    def _save_env_settings(self):
        env_lines = []
        for k, var in self.env_vars.items():
            env_lines.append(f"{k}={var.get().strip()}")

        env_path = Path(__file__).resolve().parent.parent.parent.parent / ".env"
        with open(env_path, "w", encoding="utf-8") as f:
            f.write("\n".join(env_lines) + "\n")

        Config.reload()
        self.badge_model.set_text(f"MODEL: {Config.NINEROUTER_CHAT_MODEL.split('/')[-1]}")
        messagebox.showinfo("Saved", "Configuration updated and saved to .env file successfully.")


def run_gui():
    app = PPTXJahatApp()
    app.mainloop()
