import pandas as pd
from datetime import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import time as tm  # for splash loading simulation

# --- Splash Screen ---
def show_splash(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.geometry("600x300+400+200")

    # Splash image
    img = Image.open(r"images\WhatsApp Image 2025-09-21 at 23.35.49_3f1ee5a6.jpg")
    img = img.resize((300, 150))
    logo = ImageTk.PhotoImage(img)
    logo_label = tk.Label(splash, image=logo)
    logo_label.image = logo
    logo_label.pack(pady=20)

    # Loading bar
    progress = ttk.Progressbar(splash, orient=tk.HORIZONTAL, length=400, mode='determinate')
    progress.pack(pady=20)

    for i in range(101):
        progress['value'] = i
        splash.update()
        splash.after(10)

    splash.destroy()

# --- Attendance processing ---
def process_file(file_path):
    try:
        df = pd.read_excel(file_path, engine="xlrd")
        entry_col = 'Date/Time'
        person_col = 'Name' if 'Name' in df.columns else 'Employee'
        df[entry_col] = pd.to_datetime(df[entry_col], errors='coerce')
        df['Date'] = df[entry_col].dt.date
        df['Time'] = df[entry_col].dt.time

        late_start = time(8, 15)
        comes_late_end = time(10, 0)
        early_leave_start = time(12, 0)
        early_leave_end = time(16, 30)

        df['LateStatus'] = ''
        df['EarlyLeaveStatus'] = ''
        df.loc[(df['Time'] > late_start) & (df['Time'] <= comes_late_end), 'LateStatus'] = 'Comes Late'
        df.loc[df['Time'] > comes_late_end, 'LateStatus'] = 'Very Late'
        df.loc[(df['Time'] >= early_leave_start) & (df['Time'] <= early_leave_end), 'EarlyLeaveStatus'] = 'Leave Early'

        df_agg = df.groupby([person_col, 'Date'], as_index=False).agg({
            entry_col: 'min',
            'Time': 'max',
            'LateStatus': 'first',
            'EarlyLeaveStatus': lambda x: 'Leave Early' if 'Leave Early' in x.values else ''
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

        # Clear table
        for row in tree.get_children():
            tree.delete(row)

        # Insert rows with alternating colors
        for idx, row_data in enumerate(df_agg.values.tolist()):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            tree.insert('', 'end', values=row_data, tags=(tag,))

    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- Main GUI ---
root = tk.Tk()
root.title("TDH Attendance System")
root.state('zoomed')
root.withdraw()

# Splash
show_splash(root)
root.deiconify()

# --- Top frame with logo and title ---
top_frame = tk.Frame(root, bg="#f0f0f0")
top_frame.pack(fill='x', padx=20, pady=10)

img_logo = Image.open(r"images\WhatsApp Image 2025-09-21 at 23.35.49_3f1ee5a6.jpg")
img_logo = img_logo.resize((100, 100))
logo = ImageTk.PhotoImage(img_logo)
logo_label = tk.Label(top_frame, image=logo, bg="#f0f0f0")
logo_label.pack(side='left', padx=10)

title_label = tk.Label(top_frame, text="TDH Attendance System", font=("Arial", 24, "bold"), bg="#f0f0f0")
title_label.pack(side='left', padx=20)

# --- Button frame ---
btn_frame = tk.Frame(root, bg="#f0f0f0")
btn_frame.pack(fill='x', padx=20, pady=10)

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        process_file(file_path)

load_btn = tk.Button(btn_frame, text="Load Excel File", font=("Arial", 12), bg="#F28304", fg="white",
                     padx=15, pady=8, activebackground="#F28304", command=load_file)
load_btn.pack(side='left', padx=5)

# Hover effect
def on_enter(e):
    e.widget['bg'] = '#F28304'
def on_leave(e):
    e.widget['bg'] = '#F28304'

load_btn.bind("<Enter>", on_enter)
load_btn.bind("<Leave>", on_leave)

# --- Table frame ---
table_frame = tk.Frame(root)
table_frame.pack(fill='both', expand=True, padx=20, pady=10)

columns = ['Name/Employee', 'Date', 'Time In', 'Time Out', 'LateStatus', 'EarlyLeaveStatus', 'Status']
tree = ttk.Treeview(table_frame, columns=columns, show='headings')

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#F28304", foreground="white")
style.configure("Treeview", font=("Arial", 11), rowheight=28)
tree.tag_configure('oddrow', background='#f9f9f9')
tree.tag_configure('evenrow', background='#e0e0e0')

for col in columns:
    tree.heading(col, text=col, anchor='center')
    tree.column(col, width=150, anchor='center')

scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side='right', fill='y')
tree.pack(fill='both', expand=True)

root.mainloop()
