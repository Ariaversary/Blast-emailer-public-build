import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import requests
from email import policy
from email.parser import BytesParser

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("Email Campaign Manager")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(pady=10)

        # Page 1: Subject Editor
        self.subject_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.subject_frame, text='Edit Subjects')

        self.subjects = {
            "Personal": tk.StringVar(),
            "Group": tk.StringVar(),
            "Company": tk.StringVar()
        }

        for idx, (key, var) in enumerate(self.subjects.items()):
            label = tk.Label(self.subject_frame, text=f"{key} Email Subject:")
            label.grid(row=idx, column=0, padx=10, pady=5)
            entry = tk.Entry(self.subject_frame, textvariable=var, width=50)
            entry.grid(row=idx, column=1, padx=10, pady=5)

        self.send_button = tk.Button(self.subject_frame, text="Send Emails", command=self.send_emails)
        self.send_button.grid(row=len(self.subjects), column=0, columnspan=2, pady=10)

        # Page 2: HTML File Selector
        self.html_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.html_frame, text='Select HTML Files')

        self.html_files = {}
        for key in self.subjects.keys():
            label = tk.Label(self.html_frame, text=f"{key} HTML File:")
            label.pack(pady=5)
            button = tk.Button(self.html_frame, text="Select File", command=lambda k=key: self.load_html_file(k))
            button.pack(pady=5)
            self.html_files[key] = None  # Store file paths

        # Status Frame
        self.status_frame = ttk.Frame(master)
        self.status_frame.pack(pady=10)
        self.status_label = tk.Label(self.status_frame, text="Status: Ready")
        self.status_label.pack()

    def load_html_file(self, email_type):
        file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
        if file_path:
            self.html_files[email_type] = file_path
            messagebox.showinfo("File Selected", f"{email_type} HTML file selected.")

    def send_emails(self):
        # Here you would implement the logic to send emails
        # For demonstration, we will just update the status
        self.status_label.config(text="Status: Sending emails...")
        # Simulate sending emails
        for email_type, subject in self.subjects.items():
            html_file = self.html_files[email_type]
            if html_file:
                # Simulate sending email logic
                print(f"Sending {email_type} email with subject: {subject.get()} and HTML file: {html_file}")
            else:
                print(f"⚠️ No HTML file selected for {email_type} email.")
        self.status_label.config(text="Status: Emails sent!")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
