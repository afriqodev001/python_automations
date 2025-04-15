# Streamlit App: Outlook Email Filter and Export

import streamlit as st
import win32com.client
import pythoncom
from datetime import datetime, timedelta
import os

EXPORT_PATH = os.path.join(os.getcwd(), 'EmailExports')
os.makedirs(EXPORT_PATH, exist_ok=True)

pythoncom.CoInitialize()

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)

st.title("Outlook Email Explorer & Exporter")

# Filters
filter_date = st.date_input("Select Date", datetime.now())
sender_filter = st.text_input("Filter by Sender Email")
recipient_filter = st.text_input("Filter by Recipient Email")
subject_filter = st.text_input("Filter by Subject Keyword")

# Fetch emails
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

filtered_emails = []

for message in messages:
    if message.Class == 43:  # Mail Item
        received_date = message.ReceivedTime.date()

        if received_date != filter_date:
            continue
        if sender_filter and sender_filter.lower() not in str(message.SenderEmailAddress).lower():
            continue
        if recipient_filter and recipient_filter.lower() not in str(message.To).lower():
            continue
        if subject_filter and subject_filter.lower() not in str(message.Subject).lower():
            continue

        filtered_emails.append({
            'entry_id': message.EntryID,
            'subject': message.Subject,
            'sender': f"{message.SenderName} <{message.SenderEmailAddress}>",
            'received': message.ReceivedTime,
            'body': message.Body
        })

selected_ids = []

if filtered_emails:
    st.write(f"Found {len(filtered_emails)} emails matching your filters")

    for email in filtered_emails:
        checkbox = st.checkbox(f"{email['received']} - {email['subject']} ({email['sender']})", key=email['entry_id'])
        if checkbox:
            selected_ids.append(email)
else:
    st.write("No emails found for the selected filters.")

# Export selected emails
if st.button("Export Selected Emails"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    export_file = os.path.join(EXPORT_PATH, f"Selected_Emails_{timestamp}.txt")

    with open(export_file, 'w', encoding='utf-8') as f:
        for email in selected_ids:
            f.write("=" * 50 + "\n")
            f.write(f"Subject: {email['subject']}\n")
            f.write(f"From: {email['sender']}\n")
            f.write(f"Received: {email['received']}\n\n")
            f.write(email['body'])
            f.write("\n\n")

    st.success(f"Export completed successfully! File saved at: {export_file}")
