# 🧾 Invoice PDF Extractor

A desktop GUI application that connects to your Gmail inbox, downloads invoice PDFs with a specified subject line, extracts key-value data from the PDFs, and displays the data in a table. You can also export the extracted data to a CSV file.

## 🚀 Features

- Secure Gmail login (uses app password)
- Downloads only PDF attachments from invoice emails
- Extracts key-value pairs using regex from PDF content
- Displays data in an interactive table
- Exports extracted data to `CSV`
- Threaded email fetching for UI responsiveness
- Logs activity to `invoice_app.log`

## 🛠 Tech Stack

- Python 3
- Tkinter (GUI)
- imaplib, email (email access)
- PyPDF2 (PDF reading)
- pandas (CSV export)
- threading, logging

## 📦 Installation

```bash
# Clone the repository
git clone https://github.com/your-username/invoice-pdf-extractor.git
cd invoice-pdf-extractor

# (Optional) Create and activate a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install pandas PyPDF2
 Gmail App Password Setup
Enable 2-step verification in your Google account.

Create an App Password under Google Account > Security > App Passwords

Use that app password (not your regular password) in the app.

🖥️ Usage
bash
Copy
Edit
python invoice_extractor.py
Enter your Gmail and app password.

Click "Download PDFs" to fetch matching emails.

Extracted data will be displayed in the table.

Click "Export to CSV" to save the data.

📁 Output
PDF files are saved in an attachments/ folder.

Extracted data is saved in extracted_data.csv.

📷 Screenshots
(You can insert screenshots of the UI here if you have them)

📄 License
This project is licensed under the MIT License.

🙋‍♂️ Author
Sai Krishna
GitHub: saikrishna73658
