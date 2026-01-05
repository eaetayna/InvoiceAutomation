# Invoice Automation

This project presents an invoice automation system developed in Python, focusing on identifying delinquent customers and sending personalized reminder emails via Outlook.  
I originally developed this script as a personal project for a company I worked at, saving each member of the accounts receivable team **4 to 6 hours per day**, which they previously spent performing this task manually.  
With this automation, the team now only needs to monitor customer responses, allowing them to focus on more analytical tasks.

---

## 📌 Project Overview

The automation includes:

- Reading customer and invoice data from a daily Excel database.  
- Filtering overdue invoices marked for collection.  
- Sending personalized emails to each customer via Outlook with the invoice and payment slip attached.  
- Displaying the company signature in the email (using HTML).  
- Formatting monetary values in Brazilian currency style (R$).  
- Pop-up notification upon completion.

---

## 🛠️ Libraries

- pandas
- win32com.client (pywin32)
- OpenCV (opencv-python)
- pyautogui
- datetime
- os

---

## 📊 Key Functionalities

- Automated identification of overdue invoices (daily).  
- Generation of personalized emails for each customer.  
- Integration with Microsoft Outlook for sending emails.  
- Display of the company signature image.  
- Proper formatting of dates and monetary values.  
- Notification alert when all emails have been successfully sent.

---

## 📈 Outputs

- **Emails Sent**: Personalized reminder emails are sent to all delinquent customers, including invoice number, due date, and outstanding amount.  
- **Visual Confirmation**: Optional display of the company signature image to ensure proper identification and credibility with the customer.  
- **Pop-up Notification**: A system alert confirms that all emails have been successfully sent.  
- **Formatted Data**: Monetary values are displayed in Brazilian currency format (R$) for clarity and consistency.  
- **Event Logging**: (Optional) The script can be extended to log sent emails and any errors for auditing purposes.


## 📂 Repository Structure

```text
InvoiceAutomation/
│
├── scripts/
│ └── send_invoices.py # Main automation script
├── data/
│ └── 21 - Inadimplentes.xlsx # Example invoice dataset
├── images/
│ └── signature.jpg # Company signature or logo
└── README.md # Project documentation
