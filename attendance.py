import pandas as pd
from datetime import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from PIL import Image, ImageTk

# --------- COLORS & STYLES ----------
PRIMARY_COLOR = "#F28304"
BG_COLOR = "#f5f5f5"

TITLE_FONT = ("Segoe UI", 24, "bold")
HEADING_FONT = ("Segoe UI", 12, "bold")
TEXT_FONT = ("Segoe UI", 11)
BUTTON_FONT = ("Segoe UI", 14, "bold")
STATUS_FONT = ("Segoe UI", 10)

# --- Splash Screen ---
def show_splash(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)

    w, h = 500, 280
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    splash.geometry(f"{w}x{h}+{x}+{y}")
    splash.configure(bg="white")

    # Splash image
    img = Image.open(r"images\WhatsApp Image 2025-09-21 at 23.35.49_3f1ee5a6.jpg").resize((350, 140))
    logo = ImageTk.PhotoImage(img)
    tk.Label(splash, image=logo, bg="white").pack(pady=15)
    splash.logo = logo

    # Splash title
    tk.Label(
        splash,
        text=" Welcome To The TDH Attendance System",
        font=("Segoe UI", 18, "bold"),
        bg="white",
        fg=PRIMARY_COLOR
    ).pack(pady=5)

    # --- Custom progress bar ---
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "custom.Horizontal.TProgressbar",
        troughcolor="#e0e0e0",
        background=PRIMARY_COLOR,
        thickness=22,
        bordercolor="#e0e0e0",
        lightcolor=PRIMARY_COLOR,
        darkcolor=PRIMARY_COLOR
    )

    progress = ttk.Progressbar(
        splash,
        style="custom.Horizontal.TProgressbar",
        orient=tk.HORIZONTAL,
        length=350,
        mode='determinate'
    )
    progress.pack(pady=15)

    # Animate progress
    for i in range(101):
        progress['value'] = i
        splash.update()
        splash.after(8)

    splash.destroy()

# --- Attendance processing (logic unchanged) ---
def process_file(file_path, default_type=None):
    try:
        df = pd.read_excel(file_path, engine="xlrd")
        entry_col = 'Date/Time'
        person_col = 'Name' if 'Name' in df.columns else 'Employee'
        type_col = 'EmploymentType'

        if type_col not in df.columns:
            df[type_col] = default_type

        df[entry_col] = pd.to_datetime(df[entry_col], errors='coerce')
        df['Date'] = df[entry_col].dt.date
        df['Time'] = df[entry_col].dt.time

        df['LateStatus'] = ''
        df['EarlyLeaveStatus'] = ''

        full_late_start = time(8, 30)
        full_verylate = time(10, 0)
        full_early_start = time(12, 0)
        full_early_end = time(16, 30)

        part_late_start = time(8, 30)
        part_verylate = time(9, 0)
        part_early_start = time(10, 0)
        part_early_end = time(12, 30)

        for idx, row in df.iterrows():
            emp_type = str(row[type_col]).strip().lower()
            t = row['Time']
            if pd.isnull(t):
                continue
            if emp_type == 'part-time':
                if t > part_late_start and t <= part_verylate:
                    df.at[idx, 'LateStatus'] = 'Comes Late'
                elif t > part_verylate:
                    df.at[idx, 'LateStatus'] = 'Very Late'
                if part_early_start <= t <= part_early_end:
                    df.at[idx, 'EarlyLeaveStatus'] = 'Leave Early'
            else:  # Full-time by default
                if t > full_late_start and t <= full_verylate:
                    df.at[idx, 'LateStatus'] = 'Comes Late'
                elif t > full_verylate:
                    df.at[idx, 'LateStatus'] = 'Very Late'
                if full_early_start <= t <= full_early_end:
                    df.at[idx, 'EarlyLeaveStatus'] = 'Leave Early'

        df_agg = df.groupby([person_col, 'Date'], as_index=False).agg({
            entry_col: 'min',
            'Time': 'max',
            'LateStatus': 'first',
            'EarlyLeaveStatus': lambda x: 'Leave Early' if 'Leave Early' in x.values else '',
            type_col: 'first'
        })

        df_agg['Status'] = df_agg['LateStatus'] + df_agg['EarlyLeaveStatus'].apply(lambda x: ' ' + x if x else '')
        df_agg.rename(columns={entry_col: 'Time In', 'Time': 'Time Out'}, inplace=True)
        df_agg['Time In'] = df_agg['Time In'].dt.strftime('%H:%M:%S')
        df_agg['Time Out'] = df_agg['Time Out'].apply(lambda t: t.strftime('%H:%M:%S') if pd.notnull(t) else '')

        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            df_agg.to_excel(output_file, index=False)
            messagebox.showinfo("Saved", f"Report saved as {output_file}")

        for row in tree.get_children():
            tree.delete(row)
        for idx, row_data in enumerate(df_agg.values.tolist()):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            tree.insert('', 'end', values=row_data, tags=(tag,))
    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- Main GUI ---
root = tk.Tk()
root.title("Attendance System")
root.state('zoomed')
root.configure(bg=BG_COLOR)
root.withdraw()

show_splash(root)
root.deiconify()

# --- Top frame with logo and title ---
top_frame = tk.Frame(root, bg=PRIMARY_COLOR)
top_frame.pack(fill='x')

img_logo = Image.open(r"images\WhatsApp Image 2025-09-21 at 23.35.49_3f1ee5a6.jpg").resize((180, 80))
logo = ImageTk.PhotoImage(img_logo)
logo_label = tk.Label(top_frame, image=logo, bg=PRIMARY_COLOR)
logo_label.pack(side='left', padx=10, pady=10)

title_label = tk.Label(top_frame, text="TDH Attendance System", font=TITLE_FONT, bg=PRIMARY_COLOR, fg="white")
title_label.pack(side='left', padx=20)

# --- Button frame ---
btn_frame = tk.Frame(root, bg=BG_COLOR)
btn_frame.pack(fill='x', padx=20, pady=15)

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        emp_type = None
        try:
            cols = pd.read_excel(file_path, nrows=0, engine="xlrd").columns
        except:
            cols = []

        # Validate input: only Full-Time or Part-Time allowed
        if 'EmploymentType' not in cols:
            valid = False
            while not valid:
                emp_type = simpledialog.askstring(
                    "Employee Type",
                    "Enter type for this file (Full-Time or Part-Time):"
                )
                if emp_type is None:
                    return  # User cancelled
                emp_type = emp_type.strip().title()
                if emp_type in ["Full-Time", "Part-Time"]:
                    valid = True
                else:
                    messagebox.showwarning(
                        "Invalid Input",
                        "Please enter only 'Full-Time' or 'Part-Time'."
                    )

        process_file(file_path, emp_type)

load_btn = tk.Button(btn_frame, text="📂 Load Excel File", font=BUTTON_FONT, bg=PRIMARY_COLOR, fg="white",
                     padx=20, pady=10, activebackground=PRIMARY_COLOR, cursor="hand2", command=load_file)
load_btn.pack(side='left', padx=5)

def on_enter(e): e.widget['bg'] = "#ff9933"
def on_leave(e): e.widget['bg'] = PRIMARY_COLOR
load_btn.bind("<Enter>", on_enter)
load_btn.bind("<Leave>", on_leave)

# --- Table frame ---
table_frame = tk.Frame(root, bg=BG_COLOR)
table_frame.pack(fill='both', expand=True, padx=20, pady=10)

columns = ['Name/Employee', 'Date', 'Time In', 'Time Out', 'LateStatus',
           'EarlyLeaveStatus', 'EmploymentType', 'Status']
tree = ttk.Treeview(table_frame, columns=columns, show='headings')

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview.Heading", font=HEADING_FONT, background=PRIMARY_COLOR, foreground="white")
style.configure("Treeview", font=TEXT_FONT, rowheight=28)
tree.tag_configure('oddrow', background='#ffffff')
tree.tag_configure('evenrow', background='#f1f1f1')

for col in columns:
    tree.heading(col, text=col, anchor='center')
    tree.column(col, width=150, anchor='center')

scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side='right', fill='y')
tree.pack(fill='both', expand=True)

# Status bar
status_bar = tk.Label(root, text="Load file to begin...", anchor='w', bg="#ddd", fg="#333", font=STATUS_FONT)
status_bar.pack(fill='x', side='bottom')

root.mainloop()
