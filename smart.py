# ==============================================================
# AI DATA ANALYSIS SOFTWARE
# Created by MAHRUF
#
# Modern Dynamic Dashboard
# Tkinter + Pandas + Matplotlib
#
# Features:
# - CSV / Excel file loading
# - Dynamic dataset information
# - Automatic text/numeric column detection
# - Interactive GroupBy report builder
# - Sum, Mean, Average, Max, Min, Count, Median
# - Bar, Column, Line, Pie charts
# - Report export to Excel / CSV
# - Chart export to PNG
# - Professional MAHRUF branding
# - Dynamic chart handling
# - Error handling
# - PyInstaller friendly
# ==============================================================

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ==============================================================
# APPLICATION
# ==============================================================

class MAHRUFDataAnalysisApp:

    def __init__(self, root):

        self.root = root

        # ------------------------------------------------------
        # Window settings
        # ------------------------------------------------------

        self.root.title(
            "AI Data Analysis | MAHRUF"
        )

        self.root.geometry(
            "1450x900"
        )

        self.root.minsize(
            1150,
            720
        )

        self.root.configure(
            bg="#F4F7FC"
        )

        # ------------------------------------------------------
        # Data variables
        # ------------------------------------------------------

        self.file_path = None
        self.data = None
        self.report_data = None

        self.text_columns = []
        self.numeric_columns = []

        self.figure = None
        self.canvas = None

        # ------------------------------------------------------
        # Color palette
        # ------------------------------------------------------

        self.colors = {

            "navy": "#0B1F3A",
            "navy_light": "#132F55",

            "blue": "#2185F3",
            "purple": "#7653F5",
            "green": "#15B97A",
            "orange": "#F5A623",
            "red": "#F45B69",

            "white": "#FFFFFF",
            "background": "#F4F7FC",

            "light_blue": "#EAF3FF",
            "light_purple": "#F1EDFF",
            "light_green": "#EAFBF5",
            "light_orange": "#FFF5E3",

            "text": "#172B4D",
            "muted": "#6B7A90",
            "border": "#DCE4EF"
        }

        # ------------------------------------------------------
        # Style
        # ------------------------------------------------------

        self.configure_styles()

        # ------------------------------------------------------
        # GUI
        # ------------------------------------------------------

        self.build_gui()

    # ==========================================================
    # STYLE
    # ==========================================================

    def configure_styles(self):

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "TCombobox",
            padding=7,
            font=("Segoe UI", 10)
        )

        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground=self.colors["text"],
            fieldbackground="#FFFFFF",
            rowheight=32,
            font=("Segoe UI", 9)
        )

        style.configure(
            "Treeview.Heading",
            background="#EEF3FB",
            foreground=self.colors["text"],
            font=("Segoe UI", 9, "bold"),
            padding=8
        )

        style.map(
            "Treeview",
            background=[
                ("selected", "#DCEBFF")
            ],
            foreground=[
                ("selected", self.colors["text"])
            ]
        )

    # ==========================================================
    # BUILD GUI
    # ==========================================================

    def build_gui(self):

        self.main_container = tk.Frame(
            self.root,
            bg=self.colors["background"]
        )

        self.main_container.pack(
            fill="both",
            expand=True
        )

        # Sidebar

        self.create_sidebar()

        # Content

        self.content = tk.Frame(
            self.main_container,
            bg=self.colors["background"]
        )

        self.content.pack(
            side="left",
            fill="both",
            expand=True
        )

        # Header

        self.create_header()

        # Dashboard

        self.create_scroll_area()

        # Footer

        self.create_footer()

    # ==========================================================
    # SIDEBAR
    # ==========================================================

    def create_sidebar(self):

        self.sidebar = tk.Frame(
            self.main_container,
            bg=self.colors["navy"],
            width=220
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(False)

        # ------------------------------------------------------
        # Logo
        # ------------------------------------------------------

        logo_frame = tk.Frame(
            self.sidebar,
            bg=self.colors["navy"]
        )

        logo_frame.pack(
            fill="x",
            padx=20,
            pady=(25, 25)
        )

        tk.Label(
            logo_frame,
            text="▣",
            font=("Segoe UI", 30, "bold"),
            fg="#4EA1FF",
            bg=self.colors["navy"]
        ).pack(
            side="left"
        )

        logo_text = tk.Frame(
            logo_frame,
            bg=self.colors["navy"]
        )

        logo_text.pack(
            side="left",
            padx=8
        )

        tk.Label(
            logo_text,
            text="AI DATA",
            font=("Segoe UI", 13, "bold"),
            fg="white",
            bg=self.colors["navy"]
        ).pack(
            anchor="w"
        )

        tk.Label(
            logo_text,
            text="ANALYSIS",
            font=("Segoe UI", 9),
            fg="#8EA6C7",
            bg=self.colors["navy"]
        ).pack(
            anchor="w"
        )

        # ------------------------------------------------------
        # Navigation
        # ------------------------------------------------------

        nav_items = [

            ("⌂", "Home", self.scroll_home),

            ("▣", "Load Data", self.scroll_load),

            ("◈", "Data Analysis", self.scroll_report),

            ("▤", "Reports", self.scroll_report),

            ("▥", "Charts", self.scroll_chart),

            ("⇩", "Export", self.export_report),

            ("⚙", "Settings", self.show_about)
        ]

        for icon, text, command in nav_items:

            button = tk.Button(
                self.sidebar,
                text=f"  {icon}   {text}",
                command=command,
                anchor="w",
                font=("Segoe UI", 10, "bold"),
                fg="#DCE7F7",
                bg=self.colors["navy"],
                activeforeground="white",
                activebackground="#243F68",
                bd=0,
                relief="flat",
                cursor="hand2",
                padx=18,
                pady=12
            )

            button.pack(
                fill="x",
                padx=10,
                pady=2
            )

        # ------------------------------------------------------
        # Bottom branding
        # ------------------------------------------------------

        bottom = tk.Frame(
            self.sidebar,
            bg=self.colors["navy"]
        )

        bottom.pack(
            side="bottom",
            fill="x",
            padx=20,
            pady=25
        )

        tk.Label(
            bottom,
            text="✦",
            font=("Segoe UI", 22),
            fg="#F5A623",
            bg=self.colors["navy"]
        ).pack()

        tk.Label(
            bottom,
            text="Better Data",
            font=("Segoe UI", 12, "italic"),
            fg="white",
            bg=self.colors["navy"]
        ).pack()

        tk.Label(
            bottom,
            text="Better Decisions",
            font=("Segoe UI", 11, "italic"),
            fg="#B7C6DA",
            bg=self.colors["navy"]
        ).pack()

        tk.Label(
            bottom,
            text="— MAHRUF",
            font=("Segoe UI", 9),
            fg="#6EA8FF",
            bg=self.colors["navy"]
        ).pack(
            pady=(5, 0)
        )

    # ==========================================================
    # HEADER
    # ==========================================================

    def create_header(self):

        self.header = tk.Frame(
            self.content,
            bg=self.colors["navy"],
            height=78
        )

        self.header.pack(
            fill="x"
        )

        self.header.pack_propagate(False)

        # Left title

        title_frame = tk.Frame(
            self.header,
            bg=self.colors["navy"]
        )

        title_frame.pack(
            side="left",
            padx=28
        )

        tk.Label(
            title_frame,
            text="AI Data Analysis",
            font=("Segoe UI", 21, "bold"),
            fg="white",
            bg=self.colors["navy"]
        ).pack(
            side="left"
        )

        tk.Label(
            title_frame,
            text="  ✦",
            font=("Segoe UI", 18),
            fg="#8C6CFF",
            bg=self.colors["navy"]
        ).pack(
            side="left"
        )

        tk.Label(
            title_frame,
            text="  by MAHRUF",
            font=("Segoe UI", 11, "bold"),
            fg="#2E9BFF",
            bg=self.colors["navy"]
        ).pack(
            side="left"
        )

        # Right

        right = tk.Frame(
            self.header,
            bg=self.colors["navy"]
        )

        right.pack(
            side="right",
            padx=28
        )

        tk.Label(
            right,
            text="Created by",
            font=("Segoe UI", 9),
            fg="#AABBD2",
            bg=self.colors["navy"]
        ).pack(
            side="left",
            padx=(0, 6)
        )

        tk.Label(
            right,
            text="Mahruf",
            font=("Segoe UI", 11, "bold"),
            fg="#4EA1FF",
            bg=self.colors["navy"]
        ).pack(
            side="left"
        )

    # ==========================================================
    # SCROLL AREA
    # ==========================================================

    def create_scroll_area(self):

        outer = tk.Frame(
            self.content,
            bg=self.colors["background"]
        )

        outer.pack(
            fill="both",
            expand=True
        )

        self.scroll_canvas = tk.Canvas(
            outer,
            bg=self.colors["background"],
            highlightthickness=0
        )

        scrollbar = ttk.Scrollbar(
            outer,
            orient="vertical",
            command=self.scroll_canvas.yview
        )

        self.scroll_canvas.configure(
            yscrollcommand=scrollbar.set
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.scroll_canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.dashboard = tk.Frame(
            self.scroll_canvas,
            bg=self.colors["background"]
        )

        self.dashboard_window = (
            self.scroll_canvas.create_window(
                (0, 0),
                window=self.dashboard,
                anchor="nw"
            )
        )

        self.dashboard.bind(
            "<Configure>",
            self.update_scroll_region
        )

        self.scroll_canvas.bind(
            "<Configure>",
            self.resize_dashboard
        )

        self.scroll_canvas.bind_all(
            "<MouseWheel>",
            self.mousewheel
        )

        # Build sections

        self.create_load_section()

        self.create_kpi_section()

        self.create_dataset_section()

        self.create_report_section()

        self.create_output_section()

    # ==========================================================
    # LOAD SECTION
    # ==========================================================

    def create_load_section(self):

        self.load_section = tk.Frame(
            self.dashboard,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.load_section.pack(
            fill="x",
            padx=24,
            pady=(22, 10)
        )

        tk.Label(
            self.load_section,
            text="▣   1. Load Your Data",
            font=("Segoe UI", 14, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w",
            padx=20,
            pady=(16, 2)
        )

        tk.Label(
            self.load_section,
            text="Choose a CSV or Excel file to get started with analysis.",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg="white"
        ).pack(
            anchor="w",
            padx=20
        )

        file_row = tk.Frame(
            self.load_section,
            bg="white"
        )

        file_row.pack(
            fill="x",
            padx=20,
            pady=14
        )

        self.file_entry = tk.Entry(
            file_row,
            font=("Segoe UI", 10),
            bg="#F8FAFD",
            fg=self.colors["text"],
            relief="flat",
            bd=0
        )

        self.file_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=10,
            padx=(0, 10)
        )

        self.make_button(
            file_row,
            "📁  Browse",
            self.browse_file,
            self.colors["purple"],
            13
        ).pack(
            side="right"
        )

        self.make_button(
            self.load_section,
            "▣  READ DATA",
            self.read_file,
            self.colors["green"],
            25
        ).pack(
            fill="x",
            padx=20,
            pady=(0, 18)
        )

    # ==========================================================
    # KPI CARDS
    # ==========================================================

    def create_kpi_section(self):

        self.kpi_frame = tk.Frame(
            self.dashboard,
            bg=self.colors["background"]
        )

        self.kpi_frame.pack(
            fill="x",
            padx=24,
            pady=10
        )

        self.kpi_cards = {}

        cards = [

            (
                "rows",
                "▤",
                "Total Rows",
                "0",
                self.colors["blue"]
            ),

            (
                "columns",
                "▦",
                "Total Columns",
                "0",
                self.colors["purple"]
            ),

            (
                "text",
                "≡",
                "Text Columns",
                "0",
                self.colors["green"]
            ),

            (
                "numeric",
                "#",
                "Numeric Columns",
                "0",
                self.colors["orange"]
            )
        ]

        for index, (
            key,
            icon,
            title,
            value,
            color
        ) in enumerate(cards):

            self.kpi_frame.columnconfigure(
                index,
                weight=1
            )

            card = tk.Frame(
                self.kpi_frame,
                bg=color,
                height=115
            )

            card.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=6
            )

            card.grid_propagate(False)

            tk.Label(
                card,
                text=icon,
                font=("Segoe UI", 22, "bold"),
                fg="white",
                bg=color
            ).pack(
                anchor="w",
                padx=18,
                pady=(10, 0)
            )

            tk.Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                fg="white",
                bg=color
            ).pack(
                anchor="w",
                padx=18
            )

            value_label = tk.Label(
                card,
                text=value,
                font=("Segoe UI", 20, "bold"),
                fg="white",
                bg=color
            )

            value_label.pack(
                anchor="w",
                padx=18
            )

            self.kpi_cards[key] = value_label

    # ==========================================================
    # DATASET INFORMATION
    # ==========================================================

    def create_dataset_section(self):

        wrapper = tk.Frame(
            self.dashboard,
            bg=self.colors["background"]
        )

        wrapper.pack(
            fill="x",
            padx=24,
            pady=10
        )

        wrapper.columnconfigure(
            0,
            weight=1
        )

        wrapper.columnconfigure(
            1,
            weight=1
        )

        # Dataset information

        self.info_box = tk.Frame(
            wrapper,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.info_box.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 5)
        )

        tk.Label(
            self.info_box,
            text="💡  Dataset Information",
            font=("Segoe UI", 12, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w",
            padx=18,
            pady=(15, 12)
        )

        self.dataset_info = tk.Text(
            self.info_box,
            height=7,
            font=("Segoe UI", 9),
            bg="white",
            fg=self.colors["text"],
            relief="flat",
            bd=0,
            wrap="word"
        )

        self.dataset_info.pack(
            fill="both",
            expand=True,
            padx=18,
            pady=(0, 15)
        )

        self.dataset_info.insert(
            "1.0",
            "No dataset loaded yet."
        )

        self.dataset_info.configure(
            state="disabled"
        )

        # Column detection

        self.columns_box = tk.Frame(
            wrapper,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.columns_box.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(5, 0)
        )

        tk.Label(
            self.columns_box,
            text="▱  Column Detection",
            font=("Segoe UI", 12, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w",
            padx=18,
            pady=(15, 10)
        )

        columns_inner = tk.Frame(
            self.columns_box,
            bg="white"
        )

        columns_inner.pack(
            fill="both",
            expand=True,
            padx=18,
            pady=(0, 15)
        )

        columns_inner.columnconfigure(
            0,
            weight=1
        )

        columns_inner.columnconfigure(
            1,
            weight=1
        )

        # Text columns

        text_frame = tk.Frame(
            columns_inner,
            bg=self.colors["light_green"],
            highlightbackground="#B7E8D5",
            highlightthickness=1
        )

        text_frame.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 5)
        )

        tk.Label(
            text_frame,
            text="≡  Text Columns",
            font=("Segoe UI", 10, "bold"),
            fg=self.colors["green"],
            bg=self.colors["light_green"]
        ).pack(
            anchor="w",
            padx=12,
            pady=(10, 4)
        )

        self.text_list = tk.Listbox(
            text_frame,
            height=6,
            font=("Segoe UI", 9),
            bg=self.colors["light_green"],
            fg=self.colors["text"],
            bd=0,
            highlightthickness=0
        )

        self.text_list.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 8)
        )

        # Numeric columns

        numeric_frame = tk.Frame(
            columns_inner,
            bg=self.colors["light_blue"],
            highlightbackground="#BCD9FF",
            highlightthickness=1
        )

        numeric_frame.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(5, 0)
        )

        tk.Label(
            numeric_frame,
            text="#  Numeric Columns",
            font=("Segoe UI", 10, "bold"),
            fg=self.colors["blue"],
            bg=self.colors["light_blue"]
        ).pack(
            anchor="w",
            padx=12,
            pady=(10, 4)
        )

        self.numeric_list = tk.Listbox(
            numeric_frame,
            height=6,
            font=("Segoe UI", 9),
            bg=self.colors["light_blue"],
            fg=self.colors["text"],
            bd=0,
            highlightthickness=0
        )

        self.numeric_list.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 8)
        )

    # ==========================================================
    # REPORT BUILDER
    # ==========================================================

    def create_report_section(self):

        self.report_section = tk.Frame(
            self.dashboard,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.report_section.pack(
            fill="x",
            padx=24,
            pady=10
        )

        tk.Label(
            self.report_section,
            text="▥  2. Report Builder",
            font=("Segoe UI", 13, "bold"),
            fg=self.colors["purple"],
            bg="white"
        ).pack(
            anchor="w",
            padx=18,
            pady=(15, 10)
        )

        controls = tk.Frame(
            self.report_section,
            bg="white"
        )

        controls.pack(
            fill="x",
            padx=18,
            pady=(0, 18)
        )

        # Group by

        group_frame = tk.Frame(
            controls,
            bg="white"
        )

        group_frame.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8)
        )

        tk.Label(
            group_frame,
            text="Group By Column",
            font=("Segoe UI", 9, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w"
        )

        self.group_combo = ttk.Combobox(
            group_frame,
            state="readonly",
            font=("Segoe UI", 10)
        )

        self.group_combo.pack(
            fill="x",
            pady=(5, 0)
        )

        # Aggregation

        agg_frame = tk.Frame(
            controls,
            bg="white"
        )

        agg_frame.pack(
            side="left",
            fill="x",
            expand=True,
            padx=8
        )

        tk.Label(
            agg_frame,
            text="Aggregation Method",
            font=("Segoe UI", 9, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w"
        )

        self.aggregation_combo = ttk.Combobox(
            agg_frame,
            state="readonly",
            values=[
                "sum",
                "mean",
                "average",
                "max",
                "min",
                "count",
                "median"
            ],
            font=("Segoe UI", 10)
        )

        self.aggregation_combo.pack(
            fill="x",
            pady=(5, 0)
        )

        # Value

        value_frame = tk.Frame(
            controls,
            bg="white"
        )

        value_frame.pack(
            side="left",
            fill="x",
            expand=True,
            padx=8
        )

        tk.Label(
            value_frame,
            text="Value Column",
            font=("Segoe UI", 9, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w"
        )

        self.value_combo = ttk.Combobox(
            value_frame,
            state="readonly",
            font=("Segoe UI", 10)
        )

        self.value_combo.pack(
            fill="x",
            pady=(5, 0)
        )

        # Preview

        self.make_button(
            controls,
            "⌕  Preview Report",
            self.preview_report,
            self.colors["purple"],
            18
        ).pack(
            side="left",
            padx=8,
            pady=(20, 0)
        )

        # Export

        self.make_button(
            controls,
            "⇩  Export Report",
            self.export_report,
            self.colors["blue"],
            16
        ).pack(
            side="left",
            padx=(8, 0),
            pady=(20, 0)
        )

    # ==========================================================
    # OUTPUT SECTION
    # ==========================================================

    def create_output_section(self):

        self.output_wrapper = tk.Frame(
            self.dashboard,
            bg=self.colors["background"]
        )

        self.output_wrapper.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=10
        )

        self.output_wrapper.columnconfigure(
            0,
            weight=1
        )

        self.output_wrapper.columnconfigure(
            1,
            weight=1
        )

        # ------------------------------------------------------
        # Report preview
        # ------------------------------------------------------

        self.table_box = tk.Frame(
            self.output_wrapper,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.table_box.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 6)
        )

        tk.Label(
            self.table_box,
            text="▦  Report Preview",
            font=("Segoe UI", 12, "bold"),
            fg=self.colors["text"],
            bg="white"
        ).pack(
            anchor="w",
            padx=15,
            pady=(15, 8)
        )

        table_container = tk.Frame(
            self.table_box,
            bg="white"
        )

        table_container.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(0, 12)
        )

        self.table = ttk.Treeview(
            table_container,
            show="headings"
        )

        table_scroll_y = ttk.Scrollbar(
            table_container,
            orient="vertical",
            command=self.table.yview
        )

        table_scroll_x = ttk.Scrollbar(
            table_container,
            orient="horizontal",
            command=self.table.xview
        )

        self.table.configure(
            yscrollcommand=table_scroll_y.set,
            xscrollcommand=table_scroll_x.set
        )

        self.table.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        table_scroll_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        table_scroll_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        table_container.rowconfigure(
            0,
            weight=1
        )

        table_container.columnconfigure(
            0,
            weight=1
        )

        # ------------------------------------------------------
        # Chart box
        # ------------------------------------------------------

        self.chart_box = tk.Frame(
            self.output_wrapper,
            bg="white",
            highlightbackground=self.colors["border"],
            highlightthickness=1
        )

        self.chart_box.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(6, 0)
        )

        chart_title_frame = tk.Frame(
            self.chart_box,
            bg="white"
        )

        chart_title_frame.pack(
            fill="x",
            padx=15,
            pady=(15, 5)
        )

        tk.Label(
            chart_title_frame,
            text="▥  3. Chart Builder",
            font=("Segoe UI", 12, "bold"),
            fg="#119BB9",
            bg="white"
        ).pack(
            side="left"
        )

        self.make_button(
            chart_title_frame,
            "⇩ Export Chart",
            self.export_chart,
            self.colors["blue"],
            13
        ).pack(
            side="right"
        )

        chart_controls = tk.Frame(
            self.chart_box,
            bg="white"
        )

        chart_controls.pack(
            fill="x",
            padx=15,
            pady=5
        )

        self.chart_combo = ttk.Combobox(
            chart_controls,
            state="readonly",
            values=[
                "Bar chart",
                "Column chart",
                "Line chart",
                "Pie chart"
            ],
            width=18,
            font=("Segoe UI", 9)
        )

        self.chart_combo.pack(
            side="left"
        )

        self.make_button(
            chart_controls,
            "⌕ Preview Chart",
            self.preview_chart,
            self.colors["blue"],
            15
        ).pack(
            side="left",
            padx=8
        )

        self.chart_display = tk.Frame(
            self.chart_box,
            bg="white"
        )

        self.chart_display.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        self.chart_placeholder = tk.Label(
            self.chart_display,
            text=(
                "Your chart will appear here\n\n"
                "Preview a report and select a chart type."
            ),
            font=("Segoe UI", 10),
            fg=self.colors["muted"],
            bg="white"
        )

        self.chart_placeholder.pack(
            expand=True
        )

    # ==========================================================
    # FOOTER
    # ==========================================================

    def create_footer(self):

        footer = tk.Frame(
            self.content,
            bg="white",
            height=42
        )

        footer.pack(
            fill="x"
        )

        footer.pack_propagate(False)

        self.status_label = tk.Label(
            footer,
            text="●  Ready to analyze your data",
            font=("Segoe UI", 9),
            fg=self.colors["green"],
            bg="white"
        )

        self.status_label.pack(
            side="left",
            padx=25
        )

        tk.Label(
            footer,
            text="AI Data Analysis   |   by MAHRUF",
            font=("Segoe UI", 9, "bold"),
            fg=self.colors["purple"],
            bg="white"
        ).pack(
            side="right",
            padx=25
        )

    # ==========================================================
    # BUTTON
    # ==========================================================

    def make_button(
        self,
        parent,
        text,
        command,
        bg,
        width=15
    ):

        return tk.Button(
            parent,
            text=text,
            command=command,
            width=width,
            font=("Segoe UI", 9, "bold"),
            fg="white",
            bg=bg,
            activeforeground="white",
            activebackground=bg,
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=8,
            pady=8
        )

    # ==========================================================
    # BROWSE FILE
    # ==========================================================

    def browse_file(self):

        file_path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                (
                    "CSV / Excel Files",
                    "*.csv *.xlsx *.xls"
                ),
                (
                    "CSV Files",
                    "*.csv"
                ),
                (
                    "Excel Files",
                    "*.xlsx *.xls"
                ),
                (
                    "All Files",
                    "*.*"
                )
            ]
        )

        if not file_path:
            return

        self.file_path = file_path

        self.file_entry.delete(
            0,
            tk.END
        )

        self.file_entry.insert(
            0,
            file_path
        )

        self.reset_application()

        self.set_status(
            "●  File selected. Click READ DATA to continue.",
            self.colors["blue"]
        )

    # ==========================================================
    # READ FILE
    # ==========================================================

    def read_file(self):

        if not self.file_path:

            messagebox.showerror(
                "File Required",
                "Please select a CSV or Excel file first."
            )

            return

        try:

            extension = os.path.splitext(
                self.file_path
            )[1].lower()

            # CSV

            if extension == ".csv":

                self.data = pd.read_csv(
                    self.file_path
                )

            # Excel

            elif extension in [".xlsx", ".xls"]:

                self.data = pd.read_excel(
                    self.file_path
                )

            else:

                messagebox.showerror(
                    "Unsupported File",
                    "Please select a CSV or Excel file."
                )

                return

            # Clean column names

            self.data.columns = [
                str(column).strip()
                for column in self.data.columns
            ]

            # Detect

            self.detect_columns()

            # Update dashboard

            self.update_dataset_dashboard()

            # Display data

            self.display_table(
                self.data
            )

            self.set_status(
                "●  Data loaded successfully. Ready to analyze.",
                self.colors["green"]
            )

            messagebox.showinfo(
                "Success",
                (
                    "Dataset loaded successfully!\n\n"
                    f"Rows: {len(self.data):,}\n"
                    f"Columns: {len(self.data.columns):,}"
                )
            )

        except Exception as error:

            messagebox.showerror(
                "Read Error",
                f"Unable to read the selected file.\n\n{error}"
            )

            self.set_status(
                "●  Error while reading file.",
                self.colors["red"]
            )

    # ==========================================================
    # DYNAMIC COLUMN DETECTION
    # ==========================================================

    def detect_columns(self):

        self.text_columns = []
        self.numeric_columns = []

        if self.data is None:
            return

        for column in self.data.columns:

            series = self.data[column]

            # --------------------------------------------------
            # Numeric dtype
            # --------------------------------------------------

            if pd.api.types.is_numeric_dtype(
                series
            ):

                self.numeric_columns.append(
                    column
                )

                continue

            # --------------------------------------------------
            # Numeric-looking text
            # --------------------------------------------------

            converted = pd.to_numeric(
                series,
                errors="coerce"
            )

            non_empty = series.notna().sum()

            if non_empty > 0:

                numeric_ratio = (
                    converted.notna().sum()
                    / non_empty
                )

            else:

                numeric_ratio = 0

            if numeric_ratio >= 0.80:

                self.numeric_columns.append(
                    column
                )

            else:

                self.text_columns.append(
                    column
                )

    # ==========================================================
    # UPDATE DASHBOARD
    # ==========================================================

    def update_dataset_dashboard(self):

        if self.data is None:
            return

        # ------------------------------------------------------
        # KPI
        # ------------------------------------------------------

        self.kpi_cards["rows"].config(
            text=f"{len(self.data):,}"
        )

        self.kpi_cards["columns"].config(
            text=f"{len(self.data.columns):,}"
        )

        self.kpi_cards["text"].config(
            text=f"{len(self.text_columns):,}"
        )

        self.kpi_cards["numeric"].config(
            text=f"{len(self.numeric_columns):,}"
        )

        # ------------------------------------------------------
        # Dataset info
        # ------------------------------------------------------

        filename = os.path.basename(
            self.file_path
        )

        information = (
            f"File Name   :  {filename}\n\n"
            f"Rows        :  {len(self.data):,}\n"
            f"Columns     :  {len(self.data.columns):,}\n\n"
            f"Column Headings:\n"
            f"{', '.join(map(str, self.data.columns))}"
        )

        self.dataset_info.configure(
            state="normal"
        )

        self.dataset_info.delete(
            "1.0",
            tk.END
        )

        self.dataset_info.insert(
            "1.0",
            information
        )

        self.dataset_info.configure(
            state="disabled"
        )

        # ------------------------------------------------------
        # Text list
        # ------------------------------------------------------

        self.text_list.delete(
            0,
            tk.END
        )

        for column in self.text_columns:

            self.text_list.insert(
                tk.END,
                f"• {column}"
            )

        # ------------------------------------------------------
        # Numeric list
        # ------------------------------------------------------

        self.numeric_list.delete(
            0,
            tk.END
        )

        for column in self.numeric_columns:

            self.numeric_list.insert(
                tk.END,
                f"• {column}"
            )

        # ------------------------------------------------------
        # Dropdowns
        # ------------------------------------------------------

        self.group_combo["values"] = (
            self.text_columns
        )

        self.value_combo["values"] = (
            self.numeric_columns
        )

        if self.text_columns:

            self.group_combo.current(0)

        else:

            self.group_combo.set("")

        if self.numeric_columns:

            self.value_combo.current(0)

        else:

            self.value_combo.set("")

        if self.aggregation_combo["values"]:

            self.aggregation_combo.current(0)

    # ==========================================================
    # PREVIEW REPORT
    # ==========================================================

    def preview_report(self):

        if self.data is None:

            messagebox.showerror(
                "Read Data First",
                "Please select a file and click READ DATA first."
            )

            return

        group_column = (
            self.group_combo.get()
        )

        aggregation = (
            self.aggregation_combo.get()
        )

        value_column = (
            self.value_combo.get()
        )

        # Validation

        if not group_column:

            messagebox.showerror(
                "Selection Required",
                "Please select a Group By Column."
            )

            return

        if not aggregation:

            messagebox.showerror(
                "Selection Required",
                "Please select an Aggregation Method."
            )

            return

        if not value_column:

            messagebox.showerror(
                "Selection Required",
                "Please select a Value Column."
            )

            return

        try:

            working_data = self.data.copy()

            # --------------------------------------------------
            # Convert selected value column to numeric
            # --------------------------------------------------

            working_data[value_column] = pd.to_numeric(
                working_data[value_column],
                errors="coerce"
            )

            # --------------------------------------------------
            # GroupBy
            # --------------------------------------------------

            grouped = (
                working_data
                .groupby(
                    group_column,
                    dropna=False
                )[value_column]
            )

            # --------------------------------------------------
            # Aggregation
            # --------------------------------------------------

            if aggregation == "sum":

                result = grouped.sum()

            elif aggregation in [
                "mean",
                "average"
            ]:

                result = grouped.mean()

            elif aggregation == "max":

                result = grouped.max()

            elif aggregation == "min":

                result = grouped.min()

            elif aggregation == "count":

                result = grouped.count()

            elif aggregation == "median":

                result = grouped.median()

            else:

                raise ValueError(
                    "Invalid aggregation selected."
                )

            # --------------------------------------------------
            # Convert result to DataFrame
            # --------------------------------------------------

            self.report_data = (
                result
                .reset_index()
            )

            # --------------------------------------------------
            # Rename result column
            # --------------------------------------------------

            result_name = (
                f"{aggregation.title()} "
                f"of {value_column}"
            )

            self.report_data.rename(
                columns={
                    value_column: result_name
                },
                inplace=True
            )

            # ==================================================
            # IMPORTANT:
            # Force report value to numeric
            # ==================================================

            self.report_data[result_name] = pd.to_numeric(
                self.report_data[result_name],
                errors="coerce"
            )

            # --------------------------------------------------
            # Sort descending
            # --------------------------------------------------

            self.report_data.sort_values(
                by=result_name,
                ascending=False,
                inplace=True
            )

            self.report_data.reset_index(
                drop=True,
                inplace=True
            )

            # --------------------------------------------------
            # Display report
            # --------------------------------------------------

            self.display_table(
                self.report_data
            )

            # --------------------------------------------------
            # Clear old chart
            # --------------------------------------------------

            self.clear_chart()

            self.set_status(
                "●  Report generated successfully.",
                self.colors["green"]
            )

        except Exception as error:

            messagebox.showerror(
                "Report Error",
                f"Unable to generate report.\n\n{error}"
            )

            self.set_status(
                "●  Error while generating report.",
                self.colors["red"]
            )

    # ==========================================================
    # DISPLAY TABLE
    # ==========================================================

    def display_table(
        self,
        dataframe
    ):

        # Clear rows

        for item in self.table.get_children():

            self.table.delete(
                item
            )

        # Clear columns

        self.table["columns"] = []

        if dataframe is None:
            return

        columns = list(
            dataframe.columns
        )

        self.table["columns"] = columns

        # ------------------------------------------------------
        # Headings
        # ------------------------------------------------------

        for column in columns:

            self.table.heading(
                column,
                text=str(column)
            )

            self.table.column(
                column,
                width=145,
                minwidth=100,
                anchor="center"
            )

        # ------------------------------------------------------
        # Rows
        # ------------------------------------------------------

        for _, row in dataframe.iterrows():

            values = []

            for value in row:

                if pd.isna(value):

                    values.append("")

                elif isinstance(
                    value,
                    float
                ):

                    values.append(
                        f"{value:,.2f}"
                    )

                else:

                    values.append(
                        str(value)
                    )

            self.table.insert(
                "",
                "end",
                values=values
            )

    # ==========================================================
    # EXPORT REPORT
    # ==========================================================

    def export_report(self):

        if self.data is None:

            messagebox.showerror(
                "No Data",
                "Please load and read a dataset first."
            )

            return

        if self.report_data is None:

            messagebox.showerror(
                "No Report",
                "Please preview a report before exporting."
            )

            return

        try:

            folder = os.path.dirname(
                self.file_path
            )

            base_name = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            excel_path = os.path.join(
                folder,
                f"{base_name}_analysis_report.xlsx"
            )

            csv_path = os.path.join(
                folder,
                f"{base_name}_analysis_report.csv"
            )

            # Excel

            self.report_data.to_excel(
                excel_path,
                index=False
            )

            # CSV

            self.report_data.to_csv(
                csv_path,
                index=False
            )

            self.set_status(
                "●  Report exported successfully.",
                self.colors["green"]
            )

            messagebox.showinfo(
                "Export Complete",
                (
                    "Report exported successfully!\n\n"
                    f"Excel:\n{excel_path}\n\n"
                    f"CSV:\n{csv_path}"
                )
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Unable to export report.\n\n{error}"
            )

    # ==========================================================
    # PREVIEW CHART - FIXED VERSION
    # ==========================================================

    def preview_chart(self):

        # ------------------------------------------------------
        # Check data
        # ------------------------------------------------------

        if self.data is None:

            messagebox.showerror(
                "No Data",
                "Please load and read a dataset first."
            )

            return

        # ------------------------------------------------------
        # Check report
        # ------------------------------------------------------

        if self.report_data is None:

            messagebox.showerror(
                "No Report",
                "Please preview a report first."
            )

            return

        # ------------------------------------------------------
        # Check chart type
        # ------------------------------------------------------

        chart_type = self.chart_combo.get()

        if not chart_type:

            messagebox.showerror(
                "Chart Required",
                "Please select a chart type."
            )

            return

        try:

            # ==================================================
            # Clear previous chart
            # ==================================================

            self.clear_chart()

            # Hide placeholder

            if hasattr(
                self,
                "chart_placeholder"
            ):

                self.chart_placeholder.pack_forget()

            # ==================================================
            # SAFE COPY
            # ==================================================

            chart_data = (
                self.report_data.copy()
            )

            # ==================================================
            # Identify X and Y columns
            # ==================================================

            x_column = (
                chart_data.columns[0]
            )

            y_column = (
                chart_data.columns[-1]
            )

            # ==================================================
            # IMPORTANT FIX #1
            #
            # X-axis/category values MUST be strings
            #
            # This fixes errors caused by float labels.
            # ==================================================

            chart_data[x_column] = (
                chart_data[x_column]
                .fillna("Unknown")
                .astype(str)
            )

            # ==================================================
            # IMPORTANT FIX #2
            #
            # Y-axis/value column MUST be numeric
            # ==================================================

            chart_data[y_column] = pd.to_numeric(
                chart_data[y_column],
                errors="coerce"
            )

            # ==================================================
            # Remove invalid numeric values
            # ==================================================

            chart_data = chart_data.dropna(
                subset=[y_column]
            )

            # ==================================================
            # Check chart data
            # ==================================================

            if chart_data.empty:

                messagebox.showerror(
                    "Chart Error",
                    (
                        "There is no valid numeric data "
                        "available for the chart."
                    )
                )

                return

            # ==================================================
            # Create figure
            # ==================================================

            self.figure = plt.Figure(
                figsize=(7, 4.5),
                dpi=100
            )

            ax = self.figure.add_subplot(
                111
            )

            # ==================================================
            # BAR CHART
            # ==================================================

            if chart_type == "Bar chart":

                x_values = (
                    chart_data[x_column]
                    .tolist()
                )

                y_values = (
                    chart_data[y_column]
                    .tolist()
                )

                ax.barh(
                    x_values,
                    y_values
                )

                ax.invert_yaxis()

                ax.set_xlabel(
                    str(y_column)
                )

                ax.set_ylabel(
                    str(x_column)
                )

            # ==================================================
            # COLUMN CHART
            # ==================================================

            elif chart_type == "Column chart":

                x_values = (
                    chart_data[x_column]
                    .tolist()
                )

                y_values = (
                    chart_data[y_column]
                    .tolist()
                )

                ax.bar(
                    x_values,
                    y_values
                )

                ax.set_xlabel(
                    str(x_column)
                )

                ax.set_ylabel(
                    str(y_column)
                )

                ax.tick_params(
                    axis="x",
                    rotation=40
                )

            # ==================================================
            # LINE CHART
            # ==================================================

            elif chart_type == "Line chart":

                x_values = (
                    chart_data[x_column]
                    .tolist()
                )

                y_values = (
                    chart_data[y_column]
                    .tolist()
                )

                ax.plot(
                    x_values,
                    y_values,
                    marker="o"
                )

                ax.set_xlabel(
                    str(x_column)
                )

                ax.set_ylabel(
                    str(y_column)
                )

                ax.tick_params(
                    axis="x",
                    rotation=40
                )

            # ==================================================
            # PIE CHART
            # ==================================================

            elif chart_type == "Pie chart":

                # Pie chart cannot use negative values

                pie_data = chart_data[
                    chart_data[y_column] >= 0
                ].copy()

                if pie_data.empty:

                    messagebox.showerror(
                        "Chart Error",
                        (
                            "Pie chart requires valid "
                            "non-negative numeric values."
                        )
                    )

                    return

                labels = (
                    pie_data[x_column]
                    .astype(str)
                    .tolist()
                )

                values = (
                    pie_data[y_column]
                    .tolist()
                )

                # ------------------------------------------------
                # IMPORTANT:
                # labels are explicitly strings
                # values are explicitly numeric
                # ------------------------------------------------

                ax.pie(
                    values,
                    labels=labels,
                    autopct="%1.1f%%",
                    startangle=90
                )

                ax.axis(
                    "equal"
                )

            # ==================================================
            # TITLE
            # ==================================================

            ax.set_title(
                f"{y_column} by {x_column}",
                fontsize=11,
                fontweight="bold"
            )

            # ==================================================
            # GRID
            # ==================================================

            if chart_type != "Pie chart":

                ax.grid(
                    axis="y",
                    alpha=0.20
                )

            # ==================================================
            # Layout
            # ==================================================

            self.figure.tight_layout()

            # ==================================================
            # Display inside Tkinter
            # ==================================================

            self.canvas = FigureCanvasTkAgg(
                self.figure,
                master=self.chart_display
            )

            self.canvas.draw()

            self.canvas.get_tk_widget().pack(
                fill="both",
                expand=True
            )

            self.set_status(
                "●  Chart generated successfully.",
                self.colors["green"]
            )

        except Exception as error:

            messagebox.showerror(
                "Chart Error",
                (
                    "Unable to create chart.\n\n"
                    f"{error}"
                )
            )

            self.set_status(
                "●  Error while creating chart.",
                self.colors["red"]
            )

    # ==========================================================
    # EXPORT CHART
    # ==========================================================

    def export_chart(self):

        if self.figure is None:

            messagebox.showerror(
                "No Chart",
                "Please preview a chart first."
            )

            return

        if not self.file_path:

            messagebox.showerror(
                "No File",
                "Please load a data file first."
            )

            return

        try:

            folder = os.path.dirname(
                self.file_path
            )

            base_name = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            chart_type = (
                self.chart_combo.get()
                .replace(" ", "_")
                .lower()
            )

            chart_path = os.path.join(
                folder,
                f"{base_name}_{chart_type}.png"
            )

            self.figure.savefig(
                chart_path,
                dpi=300,
                bbox_inches="tight"
            )

            self.set_status(
                "●  Chart exported successfully.",
                self.colors["green"]
            )

            messagebox.showinfo(
                "Chart Exported",
                (
                    "Chart saved successfully:\n\n"
                    f"{chart_path}"
                )
            )

        except Exception as error:

            messagebox.showerror(
                "Chart Export Error",
                f"Unable to export chart.\n\n{error}"
            )

    # ==========================================================
    # CLEAR CHART
    # ==========================================================

    def clear_chart(self):

        # ------------------------------------------------------
        # Destroy old canvas
        # ------------------------------------------------------

        if self.canvas is not None:

            try:

                self.canvas.get_tk_widget().destroy()

            except Exception:
                pass

            self.canvas = None

        # ------------------------------------------------------
        # Close old Matplotlib figure
        # ------------------------------------------------------

        if self.figure is not None:

            try:

                plt.close(
                    self.figure
                )

            except Exception:
                pass

            self.figure = None

        # ------------------------------------------------------
        # Show placeholder
        # ------------------------------------------------------

        if hasattr(
            self,
            "chart_placeholder"
        ):

            self.chart_placeholder.pack(
                expand=True
            )

    # ==========================================================
    # RESET APPLICATION
    # ==========================================================

    def reset_application(self):

        self.data = None
        self.report_data = None

        self.text_columns = []
        self.numeric_columns = []

        # ------------------------------------------------------
        # KPI reset
        # ------------------------------------------------------

        if hasattr(
            self,
            "kpi_cards"
        ):

            for key in self.kpi_cards:

                self.kpi_cards[key].config(
                    text="0"
                )

        # ------------------------------------------------------
        # Lists
        # ------------------------------------------------------

        if hasattr(
            self,
            "text_list"
        ):

            self.text_list.delete(
                0,
                tk.END
            )

        if hasattr(
            self,
            "numeric_list"
        ):

            self.numeric_list.delete(
                0,
                tk.END
            )

        # ------------------------------------------------------
        # Dataset information
        # ------------------------------------------------------

        if hasattr(
            self,
            "dataset_info"
        ):

            self.dataset_info.configure(
                state="normal"
            )

            self.dataset_info.delete(
                "1.0",
                tk.END
            )

            self.dataset_info.insert(
                "1.0",
                "No dataset loaded yet."
            )

            self.dataset_info.configure(
                state="disabled"
            )

        # ------------------------------------------------------
        # Dropdowns
        # ------------------------------------------------------

        if hasattr(
            self,
            "group_combo"
        ):

            self.group_combo.set("")
            self.value_combo.set("")
            self.aggregation_combo.set("")
            self.chart_combo.set("")

            self.group_combo["values"] = []
            self.value_combo["values"] = []

        # ------------------------------------------------------
        # Table
        # ------------------------------------------------------

        if hasattr(
            self,
            "table"
        ):

            self.display_table(
                None
            )

        # ------------------------------------------------------
        # Chart
        # ------------------------------------------------------

        self.clear_chart()

    # ==========================================================
    # STATUS
    # ==========================================================

    def set_status(
        self,
        text,
        color
    ):

        if hasattr(
            self,
            "status_label"
        ):

            self.status_label.config(
                text=text,
                fg=color
            )

    # ==========================================================
    # SCROLL FUNCTIONS
    # ==========================================================

    def update_scroll_region(
        self,
        event=None
    ):

        self.scroll_canvas.configure(
            scrollregion=(
                self.scroll_canvas.bbox("all")
            )
        )

    def resize_dashboard(
        self,
        event
    ):

        self.scroll_canvas.itemconfig(
            self.dashboard_window,
            width=event.width
        )

    def mousewheel(
        self,
        event
    ):

        try:

            self.scroll_canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )

        except Exception:
            pass

    # ==========================================================
    # NAVIGATION
    # ==========================================================

    def scroll_home(self):

        self.scroll_canvas.yview_moveto(
            0
        )

    def scroll_load(self):

        self.scroll_canvas.yview_moveto(
            0
        )

        self.file_entry.focus()

    def scroll_report(self):

        self.root.update_idletasks()

        y = self.report_section.winfo_y()

        total = self.dashboard.winfo_height()

        if total > 0:

            self.scroll_canvas.yview_moveto(
                y / total
            )

    def scroll_chart(self):

        self.root.update_idletasks()

        y = self.output_wrapper.winfo_y()

        total = self.dashboard.winfo_height()

        if total > 0:

            self.scroll_canvas.yview_moveto(
                y / total
            )

    # ==========================================================
    # ABOUT / SETTINGS
    # ==========================================================

    def show_about(self):

        messagebox.showinfo(
            "About MAHRUF",
            (
                "AI Data Analysis Software\n\n"
                "Created by Mahruf\n\n"
                "Built with:\n"
                "• Python\n"
                "• Tkinter\n"
                "• Pandas\n"
                "• Matplotlib\n\n"
                "Turn Your Data Into Meaningful Insights."
            )
        )


# ==============================================================
# MAIN
# ==============================================================

def main():

    root = tk.Tk()

    app = MAHRUFDataAnalysisApp(
        root
    )

    root.mainloop()


# ==============================================================
# START APPLICATION
# ==============================================================

if __name__ == "__main__":

    main() 