import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import imaplib
import email
from email.header import decode_header
import os
import re
import pandas as pd
from PyPDF2 import PdfReader
import getpass
import threading
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("invoice_app.log"),
        logging.StreamHandler()
    ]
)

EMAIL_FOLDER = "attachments"
SUBJECT_KEYWORD = "Invoice"
IMAP_SERVER = "imap.gmail.com"


class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice PDF Extractor")
        self.root.geometry("1000x500")
        self.invoices = []

        logging.info("Application started.")

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

        tk.Button(root, text="Export to CSV", command=self.export_csv).pack(pady=10)
        tk.Button(root, text="Exit", command=root.destroy).pack(pady=5)

    def start_email_processing(self):
        logging.info("Starting email processing in a background thread.")
        thread = threading.Thread(target=self.process_emails)
        thread.start()

    def process_emails(self):
        email_address = self.email_entry.get().strip()
        password = self.password_entry.get().strip()

        if not email_address or not password:
            messagebox.showerror("Input Error", "Please enter email and app password.")
            logging.warning("Email or password not provided.")
            return

        try:
            logging.info(f"Connecting to {IMAP_SERVER} as {email_address}.")
            imap = imaplib.IMAP4_SSL(IMAP_SERVER)
            imap.login(email_address, password)
            logging.info("Logged in successfully.")
            imap.select("inbox")
            status, messages = imap.search(None, f'SUBJECT "{SUBJECT_KEYWORD}"')

            if status != "OK":
                logging.warning("No matching emails found.")
                messagebox.showinfo("No Emails", "No matching emails found.")
                return

            os.makedirs(EMAIL_FOLDER, exist_ok=True)
            logging.info(f"Found {len(messages[0].split())} email(s) with subject '{SUBJECT_KEYWORD}'.")

            for num in messages[0].split():
                status, data = imap.fetch(num, "(RFC822)")
                raw_email = data[0][1]
                mail = email.message_from_bytes(raw_email)
                logging.info(f"Processing email: {mail.get('Subject')}")
                self.download_attachments(mail)

            imap.logout()
            logging.info("Disconnected from email server.")
            self.process_pdfs()
        except Exception as e:
            logging.error(f"Error during email processing: {e}")
            messagebox.showerror("Error", str(e))

    def download_attachments(self, mail):
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            filename = part.get_filename()
            if filename:
                filename = decode_header(filename)[0][0]
                if isinstance(filename, bytes):
                    filename = filename.decode()
                if not filename.lower().endswith(".pdf"):
                    continue
                filename = "".join(c if c.isalnum() or c in (' ', '.', '_') else "_" for c in filename)
                filepath = os.path.join(EMAIL_FOLDER, filename)
                if not os.path.isfile(filepath):
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    logging.info(f"Downloaded attachment: {filepath}")

    def extract_text_from_pdf(self, pdf_path):
        logging.info(f"Extracting text from PDF: {pdf_path}")
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
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
        logging.info(f"Extracted data: {pairs}")
        return pairs

    def process_pdfs(self):
        self.invoices.clear()
        logging.info("Processing all PDFs in the attachments folder.")
        for filename in os.listdir(EMAIL_FOLDER):
            if filename.lower().endswith(".pdf"):
                path = os.path.join(EMAIL_FOLDER, filename)
                logging.info(f"Reading PDF: {filename}")
                text = self.extract_text_from_pdf(path)
                kv = self.extract_key_value_pairs(text)
                self.invoices.append(kv)

        self.display_data()

    def display_data(self):
        logging.info("Displaying data in the UI table.")
        self.tree.delete(*self.tree.get_children())
        all_keys = set()
        for invoice in self.invoices:
            all_keys.update(invoice.keys())

        columns = list(all_keys)
        self.tree["columns"] = columns

        for col in columns:
            self.tree.heading(col, text=col)

        for invoice in self.invoices:
            row = [invoice.get(col, "") for col in columns]
            self.tree.insert("", "end", values=row)

    def export_csv(self):
        if not self.invoices:
            messagebox.showinfo("No Data", "No invoices to export.")
            logging.warning("Export attempted with no data.")
            return
        df = pd.DataFrame(self.invoices)
        df.to_csv("extracted_data.csv", index=False)
        logging.info("Exported data to 'extracted_data.csv'.")
        messagebox.showinfo("Success", "Data exported to 'extracted_data.csv'")


if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
