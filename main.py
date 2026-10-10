import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pandas as pd
import pymupdf
import os


# -----------------------------
# Main Window
# -----------------------------

window = tk.Tk()
window.title("File Converter")
window.geometry("700x550")
window.resizable(False, False)


# -----------------------------
# Variables
# -----------------------------

selected_file = tk.StringVar()
output_format = tk.StringVar(value="Select Format")


# -----------------------------
# Select File
# -----------------------------

def select_file():

    file_path = filedialog.askopenfilename(
        title="Select a file"
    )

    if file_path:

        selected_file.set(file_path)

        file_name = os.path.basename(file_path)

        file_label.config(
            text=f"Selected: {file_name}"
        )


# -----------------------------
# Convert Image
# -----------------------------

def convert_image(file_path, output_path):

    image = Image.open(file_path)

    image = image.convert("RGB")

    image.save(output_path)


# -----------------------------
# Convert TXT to PDF
# -----------------------------

def txt_to_pdf(file_path, output_path):

    pdf = canvas.Canvas(
        output_path,
        pagesize=letter
    )

    width, height = letter

    y = height - 50

    with open(
    file_path,
    "r",
    encoding="utf-8-sig"
) as file:

     lines = file.readlines()

    for line in lines:

        pdf.drawString(
            50,
            y,
            line.strip()
        )

        y -= 15

        if y < 50:

            pdf.showPage()

            y = height - 50

    pdf.save()


# -----------------------------
# Convert CSV to Excel
# -----------------------------

def csv_to_excel(file_path, output_path):

    data = pd.read_csv(file_path)

    data.to_excel(
        output_path,
        index=False
    )


# -----------------------------
# Convert Excel to CSV
# -----------------------------

def excel_to_csv(file_path, output_path):

    data = pd.read_excel(file_path)

    data.to_csv(
        output_path,
        index=False
    )


# -----------------------------
# Convert PDF to TXT
# -----------------------------
def pdf_to_txt(file_path, output_path):

    document = pymupdf.open(file_path)

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as file:

        for page in document:

            text = page.get_text()

            file.write(text)

    document.close()


# -----------------------------
# Main Conversion Function
# -----------------------------

def convert_file():

    file_path = selected_file.get()

    format_choice = output_format.get()


    if not file_path:

        messagebox.showerror(
            "Error",
            "Please select a file first."
        )

        return


    if format_choice == "Select Format":

        messagebox.showerror(
            "Error",
            "Please select an output format."
        )

        return


    file_name = os.path.basename(file_path)

    name_without_extension = os.path.splitext(
        file_name
    )[0]


    # -------------------------
    # Image conversion
    # -------------------------

    if format_choice == "JPG":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".jpg",
            initialfile=name_without_extension + ".jpg",
            filetypes=[
                ("JPG files", "*.jpg")
            ]
        )

        if output_path:

            convert_image(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "File converted successfully!"
            )


    elif format_choice == "PNG":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=name_without_extension + ".png",
            filetypes=[
                ("PNG files", "*.png")
            ]
        )

        if output_path:

            convert_image(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "File converted successfully!"
            )


    # -------------------------
    # TXT → PDF
    # -------------------------

    elif format_choice == "PDF":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=name_without_extension + ".pdf",
            filetypes=[
                ("PDF files", "*.pdf")
            ]
        )

        if output_path:

            txt_to_pdf(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "PDF created successfully!"
            )


    # -------------------------
    # CSV → Excel
    # -------------------------

    elif format_choice == "XLSX":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=name_without_extension + ".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx")
            ]
        )

        if output_path:

            csv_to_excel(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "Excel file created successfully!"
            )


    # -------------------------
    # Excel → CSV
    # -------------------------

    elif format_choice == "CSV":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=name_without_extension + ".csv",
            filetypes=[
                ("CSV files", "*.csv")
            ]
        )

        if output_path:

            excel_to_csv(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "CSV file created successfully!"
            )


    # -------------------------
    # PDF → TXT
    # -------------------------

    elif format_choice == "TXT":

        output_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=name_without_extension + ".txt",
            filetypes=[
                ("Text files", "*.txt")
            ]
        )

        if output_path:

            pdf_to_txt(
                file_path,
                output_path
            )

            messagebox.showinfo(
                "Success",
                "Text file created successfully!"
            )

# -----------------------------
# GUI Design
# -----------------------------

# Window
window.title("FileX - File Converter")
window.geometry("1100x700")
window.resizable(True, True)

# -----------------------------
# Colors
# -----------------------------

BG_COLOR = "#151218"
SIDEBAR_COLOR = "#211C24"
CARD_COLOR = "#242027"
CARD_COLOR_2 = "#2B2530"
TEXT_COLOR = "#F5F0F5"
SECONDARY_TEXT = "#AAA2AE"
ACCENT_COLOR = "#D89BE8"
ACCENT_DARK = "#9C6BAA"
BORDER_COLOR = "#413844"
SUCCESS_COLOR = "#72D6A0"


# -----------------------------
# Configure Window
# -----------------------------

window.configure(bg=BG_COLOR)


# -----------------------------
# Main Container
# -----------------------------

main_container = tk.Frame(
    window,
    bg=BG_COLOR
)

main_container.pack(
    fill="both",
    expand=True
)


# ============================================================
# LEFT SIDEBAR
# ============================================================

sidebar = tk.Frame(
    main_container,
    bg=SIDEBAR_COLOR,
    width=230
)

sidebar.pack(
    side="left",
    fill="y"
)

sidebar.pack_propagate(False)


# -----------------------------
# Logo
# -----------------------------

logo = tk.Label(
    sidebar,
    text="F  FILEX",
    font=("Arial", 20, "bold"),
    fg=TEXT_COLOR,
    bg=SIDEBAR_COLOR
)

logo.pack(
    pady=(35, 50),
    padx=25,
    anchor="w"
)


# -----------------------------
# Sidebar Button Function
# -----------------------------

def sidebar_button(text):

    button = tk.Button(
        sidebar,
        text=text,
        font=("Arial", 11),
        fg=SECONDARY_TEXT,
        bg=SIDEBAR_COLOR,
        activeforeground=TEXT_COLOR,
        activebackground=CARD_COLOR,
        bd=0,
        relief="flat",
        anchor="w",
        padx=25,
        width=20,
        height=2,
        cursor="hand2"
    )

    button.pack(
        pady=3
    )

    return button


# -----------------------------
# Sidebar Navigation
# -----------------------------

dashboard_btn = sidebar_button("◉   Dashboard")

converter_btn = sidebar_button("▣   Converter")

history_btn = sidebar_button("◷   History")


# -----------------------------
# Bottom Sidebar
# -----------------------------

sidebar_bottom = tk.Frame(
    sidebar,
    bg=SIDEBAR_COLOR
)

sidebar_bottom.pack(
    side="bottom",
    pady=25,
    fill="x"
)


settings_btn = sidebar_button("⚙   Settings")

support_btn = sidebar_button("?   Support")


# ============================================================
# RIGHT SIDE
# ============================================================

content = tk.Frame(
    main_container,
    bg=BG_COLOR
)

content.pack(
    side="right",
    fill="both",
    expand=True
)


# ============================================================
# TOP BAR
# ============================================================

top_bar = tk.Frame(
    content,
    bg=BG_COLOR,
    height=75
)

top_bar.pack(
    fill="x",
    padx=35
)

top_bar.pack_propagate(False)


# -----------------------------
# Top Title
# -----------------------------

top_title = tk.Label(
    top_bar,
    text="File Converter",
    font=("Arial", 13, "bold"),
    fg=TEXT_COLOR,
    bg=BG_COLOR
)

top_title.pack(
    side="left",
    pady=25
)


# -----------------------------
# Status
# -----------------------------

status_top = tk.Label(
    top_bar,
    text="●  READY",
    font=("Arial", 10, "bold"),
    fg=SUCCESS_COLOR,
    bg=BG_COLOR
)

status_top.pack(
    side="right",
    pady=25
)


# ============================================================
# WELCOME SECTION
# ============================================================

welcome_frame = tk.Frame(
    content,
    bg=BG_COLOR
)

welcome_frame.pack(
    fill="x",
    padx=35,
    pady=(15, 20)
)


welcome_title = tk.Label(
    welcome_frame,
    text="Welcome to FileX",
    font=("Arial", 25, "bold"),
    fg=TEXT_COLOR,
    bg=BG_COLOR
)

welcome_title.pack(
    anchor="w"
)


welcome_text = tk.Label(
    welcome_frame,
    text="Convert your files quickly, easily and securely.",
    font=("Arial", 11),
    fg=SECONDARY_TEXT,
    bg=BG_COLOR
)

welcome_text.pack(
    anchor="w",
    pady=(6, 0)
)


# ============================================================
# MAIN CARDS
# ============================================================

cards_frame = tk.Frame(
    content,
    bg=BG_COLOR
)

cards_frame.pack(
    fill="x",
    padx=35
)


# ============================================================
# SELECT FILE CARD
# ============================================================

file_card = tk.Frame(
    cards_frame,
    bg=CARD_COLOR,
    width=350,
    height=220
)

file_card.pack(
    side="left",
    fill="both",
    expand=True,
    padx=(0, 10)
)

file_card.pack_propagate(False)


# Card heading

file_card_title = tk.Label(
    file_card,
    text="SELECT FILE",
    font=("Arial", 10, "bold"),
    fg=SECONDARY_TEXT,
    bg=CARD_COLOR
)

file_card_title.pack(
    anchor="w",
    padx=25,
    pady=(22, 10)
)


file_card_heading = tk.Label(
    file_card,
    text="Choose a file",
    font=("Arial", 19, "bold"),
    fg=TEXT_COLOR,
    bg=CARD_COLOR
)

file_card_heading.pack(
    anchor="w",
    padx=25
)


file_card_description = tk.Label(
    file_card,
    text="Select the file you want to convert.",
    font=("Arial", 10),
    fg=SECONDARY_TEXT,
    bg=CARD_COLOR
)

file_card_description.pack(
    anchor="w",
    padx=25,
    pady=(5, 15)
)


# -----------------------------
# Browse Button
# -----------------------------

select_button = tk.Button(
    file_card,
    text="＋  Browse File",
    font=("Arial", 11, "bold"),
    fg=TEXT_COLOR,
    bg=ACCENT_DARK,
    activeforeground=TEXT_COLOR,
    activebackground=ACCENT_COLOR,
    bd=0,
    relief="flat",
    cursor="hand2",
    command=select_file
)

select_button.pack(
    padx=25,
    anchor="w",
    ipadx=15,
    ipady=7
)


# ============================================================
# CONVERSION CARD
# ============================================================

conversion_card = tk.Frame(
    cards_frame,
    bg=CARD_COLOR,
    width=350,
    height=220
)

conversion_card.pack(
    side="right",
    fill="both",
    expand=True,
    padx=(10, 0)
)

conversion_card.pack_propagate(False)


conversion_title = tk.Label(
    conversion_card,
    text="CONVERSION",
    font=("Arial", 10, "bold"),
    fg=SECONDARY_TEXT,
    bg=CARD_COLOR
)

conversion_title.pack(
    anchor="w",
    padx=25,
    pady=(22, 10)
)


format_label = tk.Label(
    conversion_card,
    text="Output Format",
    font=("Arial", 12, "bold"),
    fg=TEXT_COLOR,
    bg=CARD_COLOR
)

format_label.pack(
    anchor="w",
    padx=25,
    pady=(2, 8)
)


# -----------------------------
# Format Dropdown
# -----------------------------

style = ttk.Style()

style.theme_use("clam")

style.configure(
    "Custom.TCombobox",
    fieldbackground=CARD_COLOR_2,
    background=CARD_COLOR_2,
    foreground=TEXT_COLOR,
    arrowcolor=ACCENT_COLOR,
    bordercolor=BORDER_COLOR,
    lightcolor=BORDER_COLOR,
    darkcolor=BORDER_COLOR,
    padding=8
)

style.map(
    "Custom.TCombobox",
    fieldbackground=[
        ("readonly", CARD_COLOR_2)
    ],
    foreground=[
        ("readonly", TEXT_COLOR)
    ]
)


format_dropdown = ttk.Combobox(
    conversion_card,
    textvariable=output_format,
    values=[
        "JPG",
        "PNG",
        "PDF",
        "XLSX",
        "CSV",
        "TXT"
    ],
    state="readonly",
    width=24,
    style="Custom.TCombobox"
)

format_dropdown.pack(
    padx=25,
    anchor="w"
)


# -----------------------------
# Convert Button
# -----------------------------

convert_button = tk.Button(
    conversion_card,
    text="CONVERT  →",
    font=("Arial", 11, "bold"),
    fg=TEXT_COLOR,
    bg=ACCENT_DARK,
    activeforeground=TEXT_COLOR,
    activebackground=ACCENT_COLOR,
    bd=0,
    relief="flat",
    cursor="hand2",
    command=convert_file
)

convert_button.pack(
    padx=25,
    pady=18,
    anchor="w",
    ipadx=25,
    ipady=7
)


# ============================================================
# SELECTED FILE CARD
# ============================================================

selected_card = tk.Frame(
    content,
    bg=CARD_COLOR,
    height=125
)

selected_card.pack(
    fill="x",
    padx=35,
    pady=20
)

selected_card.pack_propagate(False)


selected_title = tk.Label(
    selected_card,
    text="SELECTED FILE",
    font=("Arial", 10, "bold"),
    fg=SECONDARY_TEXT,
    bg=CARD_COLOR
)

selected_title.pack(
    anchor="w",
    padx=25,
    pady=(18, 3)
)


# IMPORTANT:
# Keep the variable name "file_label"
# because select_file() already uses it.

file_label = tk.Label(
    selected_card,
    text="No file selected",
    font=("Arial", 12),
    fg=TEXT_COLOR,
    bg=CARD_COLOR,
    anchor="w"
)

file_label.pack(
    fill="x",
    padx=25,
    pady=8
)


# ============================================================
# BOTTOM STATUS CARD
# ============================================================

status_card = tk.Frame(
    content,
    bg=CARD_COLOR_2,
    height=85
)

status_card.pack(
    fill="x",
    padx=35,
    pady=(0, 15)
)

status_card.pack_propagate(False)


status_title = tk.Label(
    status_card,
    text="CONVERSION STATUS",
    font=("Arial", 9, "bold"),
    fg=SECONDARY_TEXT,
    bg=CARD_COLOR_2
)

status_title.pack(
    anchor="w",
    padx=25,
    pady=(13, 2)
)


status_label = tk.Label(
    status_card,
    text="●  Ready to convert",
    font=("Arial", 11, "bold"),
    fg=SUCCESS_COLOR,
    bg=CARD_COLOR_2
)

status_label.pack(
    anchor="w",
    padx=25
)


# ============================================================
# FOOTER
# ============================================================

footer = tk.Label(
    content,
    text="FileX  •  Python File Converter",
    font=("Arial", 9),
    fg=SECONDARY_TEXT,
    bg=BG_COLOR
)

footer.pack(
    side="bottom",
    pady=10
)


# -----------------------------
# Start GUI
# -----------------------------

window.mainloop()