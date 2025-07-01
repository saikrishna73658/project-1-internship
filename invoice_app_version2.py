import tkinter as tk
from tkinter import ttk, messagebox
import imaplib
import email
from email.header import decode_header
import os
import getpass
import pdfplumber
import re
import pandas as pd
import threading
import logging

IMAP_SERVER = "imap.gmail.com"
EMAIL_FOLDER = "attachments"
SUBJECT_KEYWORD = "Invoice"
CSV_OUTPUT = "extracted_data.csv"
LOG_FILE = "invoice_app.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice PDF Extractor")
        self.root.geometry("1000x500")
        self.invoices = []

        frame_login = tk.LabelFrame(root, text="Email Login")
        frame_login.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_login, text="Email:").grid(row=0, column=0, padx=5, pady=5)
        self.email_entry = tk.Entry(frame_login, width=30)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_login, text="App Password:").grid(row=0, column=2, padx=5, pady=5)
        self.password_entry = tk.Entry(frame_login, width=30, show="*")
        self.password_entry.grid(row=0, column=3, padx=5, pady=5)

        tk.Button(frame_login, text="Download PDFs", command=self.start_email_processing).grid(row=0, column=4, padx=10)

        self.tree = ttk.Treeview(root, show='headings')
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

        tk.Button(root, text="Export to CSV", command=self.export_csv).pack(pady=5)
        tk.Button(root, text="Exit", command=root.destroy).pack(pady=5)

    def start_email_processing(self):
        thread = threading.Thread(target=self.process_emails)
        thread.start()

    def process_emails(self):
        email_address = self.email_entry.get().strip()
        password = self.password_entry.get().strip()

        if not email_address or not password:
            messagebox.showerror("Input Error", "Please enter email and app password.")
            return

        try:
            imap = imaplib.IMAP4_SSL(IMAP_SERVER)
            imap.login(email_address, password)
            logging.info("Login successful.")
            imap.select("inbox")

            status, messages = imap.search(None, f'SUBJECT "{SUBJECT_KEYWORD}"')

            if status != "OK" or not messages[0]:
                messagebox.showinfo("No Emails", "No matching emails found.")
                return

            os.makedirs(EMAIL_FOLDER, exist_ok=True)

            for num in messages[0].split():
                status, data = imap.fetch(num, "(RFC822)")
                raw_email = data[0][1]
                mail = email.message_from_bytes(raw_email)
                self.download_attachments(mail)

            imap.logout()
            self.process_pdfs()

        except Exception as e:
            logging.error(f"Error: {e}")
            messagebox.showerror("Error", str(e))

    def decode_filename(self, raw_filename):
        decoded_parts = decode_header(raw_filename)
        return ''.join([
            part.decode(enc or 'utf-8') if isinstance(part, bytes) else part
            for part, enc in decoded_parts
        ])

    def download_attachments(self, mail):
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename:
                filename = self.decode_filename(filename)
                if not filename.lower().endswith(".pdf"):
                    continue

                filename = "".join(c if c.isalnum() or c in (' ', '.', '_') else '_' for c in filename)
                filepath = os.path.join(EMAIL_FOLDER, filename)

                if not os.path.isfile(filepath):
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    logging.info(f"Downloaded: {filepath}")

    def extract_text_from_pdf(self, pdf_path):
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            txt_path = os.path.splitext(pdf_path)[0] + ".txt"
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception as e:
            logging.warning(f"Failed to read {pdf_path}: {e}")
        return text

    def extract_key_value_pairs(self, text):
        pairs = {}
        lines = [line.strip() for line in text.split("\n") if ":" in line]
        for line in lines:
            if re.match(r".+?:\s?.+", line):
                key, value = line.split(":", 1)
                key_clean = re.sub(r"[^A-Za-z0-9 ]+", "", key).strip()
                value_clean = value.strip()
                if key_clean:
                    pairs[key_clean] = value_clean
        return pairs

    def process_pdfs(self):
        self.invoices.clear()
        for filename in os.listdir(EMAIL_FOLDER):
            if filename.lower().endswith(".pdf"):
                path = os.path.join(EMAIL_FOLDER, filename)
                text = self.extract_text_from_pdf(path)
                kv = self.extract_key_value_pairs(text)
                if kv:
                    self.invoices.append(kv)
        self.display_data()

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        all_keys = set()
        for invoice in self.invoices:
            all_keys.update(invoice.keys())

        columns = list(all_keys)
        self.tree["columns"] = columns

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        for invoice in self.invoices:
            row = [invoice.get(col, "") for col in columns]
            self.tree.insert("", "end", values=row)

    def export_csv(self):
        if not self.invoices:
            messagebox.showinfo("No Data", "No invoices to export.")
            return
        df = pd.DataFrame(self.invoices)
        df.to_csv(CSV_OUTPUT, index=False)
        messagebox.showinfo("Success", f"Data exported to '{CSV_OUTPUT}'")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
