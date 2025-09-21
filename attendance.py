import pandas as pd
from datetime import time
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
import smtplib
from email.message import EmailMessage

# --- Function to send email ---
def send_email(file_path, recipient_email):
    try:
        # --- Email configuration ---
        sender_email = "sultanimahfoza@gmail.com"   # Replace with your email
        sender_password = "12345"         # Replace with your email password or app password

        # Create email
        msg = EmailMessage()
        msg['Subject'] = "Attendance Report"
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content("Dear recipient,\n\nPlease find attached the attendance report.\n\nBest regards.")

        # Attach the Excel file
        with open(file_path, 'rb') as f:
            file_data = f.read()
            file_name = file_path.split('/')[-1]
        msg.add_attachment(file_data, maintype='application',
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           filename=file_name)

        # Send email via Gmail SMTP (adjust server for Outlook etc.)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)

        messagebox.showinfo("Success", f"Email sent successfully to {recipient_email}")

    except Exception as e:
        messagebox.showerror("Email Error", str(e))

# --- Function to process attendance ---
def process_file(file_path):
    try:
        df = pd.read_excel(file_path, engine="xlrd")
        
        entry_col = 'Date/Time'
        person_col = 'Name' if 'Name' in df.columns else 'Employee'
        
        df[entry_col] = pd.to_datetime(df[entry_col], errors='coerce')
        df['Date'] = df[entry_col].dt.date
        df['Time'] = df[entry_col].dt.time
        
        # Thresholds
        late_start = time(8, 15)
        comes_late_end = time(10, 0)
        early_leave_start = time(12, 0)
        early_leave_end = time(16, 30)
        
        # Initialize status columns
        df['LateStatus'] = ''
        df['EarlyLeaveStatus'] = ''
        
        # Late logic
        df.loc[(df['Time'] > late_start) & (df['Time'] <= comes_late_end), 'LateStatus'] = 'Comes Late'
        df.loc[df['Time'] > comes_late_end, 'LateStatus'] = 'Very Late'
        
        # Early leave logic
        df.loc[(df['Time'] >= early_leave_start) & (df['Time'] <= early_leave_end), 'EarlyLeaveStatus'] = 'Leave Early'
        
        # Aggregate per person and date
        df_agg = df.groupby([person_col, 'Date'], as_index=False).agg({
            entry_col: 'min',
            'Time': 'max',
            'LateStatus': 'first',
            'EarlyLeaveStatus': lambda x: 'Leave Early' if 'Leave Early' in x.values else ''
        })
        
        # Combine status
        df_agg['Status'] = df_agg['LateStatus'] + df_agg['EarlyLeaveStatus'].apply(lambda x: ' ' + x if x else '')
        
        # Rename columns
        df_agg.rename(columns={entry_col: 'Time In', 'Time': 'Time Out'}, inplace=True)
        
        # Format time
        df_agg['Time In'] = df_agg['Time In'].dt.strftime('%H:%M:%S')
        df_agg['Time Out'] = df_agg['Time Out'].apply(lambda t: t.strftime('%H:%M:%S') if pd.notnull(t) else '')
        
        # Save Excel
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            df_agg.to_excel(output_file, index=False)
            messagebox.showinfo("Saved", f"Report saved as {output_file}")

            # Ask if user wants to send email
            if messagebox.askyesno("Send Email", "Do you want to send this report by email?"):
                recipient = simpledialog.askstring("Recipient Email", "Enter recipient email:")
                if recipient:
                    send_email(output_file, recipient)
        
        # Display in table
        for row in tree.get_children():
            tree.delete(row)
        for _, row in df_agg.iterrows():
            tree.insert('', 'end', values=list(row))
        
    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- GUI ---
root = tk.Tk()

root.title("TDH Attendance System")
# root.geometry("1400x720")
root.state('zoomed')

# Buttons
btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        process_file(file_path)

load_btn = tk.Button(btn_frame, text="Load Excel File", command=load_file)
load_btn.pack()

# Table to display result
columns = ['Name/Employee', 'Date', 'Time In', 'Time Out', 'LateStatus', 'EarlyLeaveStatus', 'Status']
tree = ttk.Treeview(root, columns=columns, show='headings')
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)
tree.pack(expand=True, fill='both', pady=10)

root.mainloop()
