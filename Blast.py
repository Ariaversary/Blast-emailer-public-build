import extract_msg
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.simpledialog as simpledialog
from tkinter import scrolledtext
import re
import base64
import os
import pandas as pd
import requests
import time
import threading
import json

# Load credentials from config.json
try:
    with open("config.json", "r") as f:
        config = json.load(f)
except FileNotFoundError:
    raise FileNotFoundError("❌ 'config.json' file not found. Please create one with the required credentials.")
except json.JSONDecodeError as e:
    raise ValueError(f"❌ Invalid JSON format in config.json: {e}")

# Validate expected keys
required_keys = ["tenant_id", "client_id", "client_secret", "sender_email"]
for key in required_keys:
    if key not in config or not config[key].strip():
        raise ValueError(f"❌ Missing or empty '{key}' in config.json")

# Set variables
TENANT_ID = config["tenant_id"]
CLIENT_ID = config["client_id"]
CLIENT_SECRET = config["client_secret"]
SENDER_EMAIL = config["sender_email"]



cached_token = None
token_acquired_time = 0


# ────────────────────────────────
# Microsoft Graph API credentials
# ────────────────────────────────

def get_access_token():
    global cached_token, token_acquired_time
    now = time.time()
    if cached_token is None or now - token_acquired_time > 3000:
        token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        response = requests.post(token_url, data=data)
        if response.status_code == 200:
            cached_token = response.json()["access_token"]
            token_acquired_time = now
        else:
            raise Exception(f"Failed to get access token: {response.text}")
    return cached_token


def extract_html_and_image_map(msg_path):
    msg = extract_msg.Message(msg_path)
    html_body = msg.htmlBody or ""
    if isinstance(html_body, bytes):
        html_body = html_body.decode(errors="replace")

    cid_matches = re.findall(r'cid:([^"\'>]+)', html_body)
    cid_to_file = {}

    for cid in sorted(set(cid_matches)):
        file_path = filedialog.askopenfilename(
            title=f"Select image for CID: {cid}",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.webp")]
        )
        if file_path:
            cid_to_file[cid] = file_path
        else:
            messagebox.showwarning("Missing Image", f"No image selected for CID: {cid}")
    return msg.subject or "email", html_body, cid_to_file


def prepare_inline_attachments(cid_to_file):
    attachments = []
    for cid, file_path in cid_to_file.items():
        with open(file_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(file_path),
            "contentId": cid,
            "isInline": True,
            "contentBytes": encoded
        }
        attachments.append(attachment)
    return attachments


def send_email(to_email, subject, html_body, token, attachments, cc_emails=None):
    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": to_email}}],
        "attachments": attachments
    }

    if cc_emails:
        message["ccRecipients"] = [{"emailAddress": {"address": cc}} for cc in cc_emails]

    response = requests.post(url, headers=headers, json={"message": message})
    return response.status_code == 202, response.text


def run():
    root = tk.Tk()
    root.title("Bulk Email Sender")

    # ──────────────
    # Log box
    # ──────────────
    status_log = scrolledtext.ScrolledText(root, width=90, height=20, state='disabled', wrap='word')
    status_log.pack(padx=10, pady=10)

    def log(msg):
        def append():
            status_log.config(state='normal')
            status_log.insert(tk.END, msg + "\n")
            status_log.see(tk.END)
            status_log.config(state='disabled')
        root.after(0, append)

    # ──────────────
    # Input phase
    # ──────────────
    msg_path = filedialog.askopenfilename(
        title="Select MSG File",
        filetypes=[("MSG Files", "*.msg")]
    )
    if not msg_path:
        return

    excel_path = filedialog.askopenfilename(
        title="Select Excel Sheet with Email Column",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_path:
        return

    try:
        subject, html_body, cid_to_file = extract_html_and_image_map(msg_path)
        attachments = prepare_inline_attachments(cid_to_file)

        xlsx = pd.ExcelFile(excel_path)
        sheet_name = simpledialog.askstring("Sheet Selection",
            f"Available sheets:\n{', '.join(xlsx.sheet_names)}\n\nEnter sheet name to use:")
        if sheet_name not in xlsx.sheet_names:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' not found.")
            return

        df = pd.read_excel(xlsx, sheet_name=sheet_name)
        email_col = next((col for col in df.columns if str(col).strip().lower() == "email"), None)
        if not email_col:
            messagebox.showerror("Error", "No 'Email' column found in the Excel file.")
            return

    except Exception as setup_error:
        messagebox.showerror("Setup Error", str(setup_error))
        return

    # ──────────────
    # Delay & CC Input
    # ──────────────
    delay_var = tk.DoubleVar(value=2.0)
    cc_var = tk.StringVar()

    input_win = tk.Toplevel(root)
    input_win.title("Email Settings")

    tk.Label(input_win, text="Delay between emails (seconds):").pack(pady=(10, 2))
    tk.Entry(input_win, textvariable=delay_var, width=10).pack()

    tk.Label(input_win, text="CC (comma-separated emails):").pack(pady=(10, 2))
    tk.Entry(input_win, textvariable=cc_var, width=50).pack()

    def send_all_emails():
        input_win.destroy()
        cc_list = [cc.strip() for cc in cc_var.get().split(",") if cc.strip()]
        delay = delay_var.get()
        log("🔄 Starting email send process...")

        def thread_func():
            try:
                token = get_access_token()
                for idx, row in df.iterrows():
                    email = str(row[email_col]).strip()
                    if email:
                        success, result = send_email(email, subject, html_body, token, attachments, cc_emails=cc_list)
                        if success:
                            log(f"✅ Sent to {email}")
                        else:
                            log(f"❌ Failed to {email}: {result}")
                        time.sleep(delay)
                root.after(0, lambda: messagebox.showinfo("Done", "All emails sent."))
            except Exception as e:
                log(f"❌ Error: {e}")
                root.after(0, lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=thread_func, daemon=True).start()

    tk.Button(input_win, text="Start Sending", command=send_all_emails).pack(pady=10)
    input_win.grab_set()

    root.mainloop()


if __name__ == "__main__":
    run()
