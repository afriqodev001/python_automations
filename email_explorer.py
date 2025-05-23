# Streamlit Outlook Email Explorer - Project Structure

sender = f"{mail.SenderName} <{mail.SenderEmailAddress}>" if mail.SenderEmailAddress else mail.SenderName

# main.py
import streamlit as st
from email_client import OutlookClient
from chain_handler import group_by_conversation
from exporter import export_email, export_conversation
from logger import get_logger

logger = get_logger()

st.title("Outlook Email Explorer")

keyword = st.text_input("Search Keyword")

if keyword:
    client = OutlookClient()
    emails = client.search_emails(keyword)
    grouped = group_by_conversation(emails)

    st.write(f"Found {len(emails)} matching emails in {len(grouped)} conversations")

    for cid, chain in grouped.items():
        with st.expander(f"Conversation: {chain[0].ConversationTopic} ({len(chain)} Emails)"):
            for mail in chain:
                st.markdown(f"### {mail.Subject}")
                st.write(f"From: {mail.SenderEmailAddress}")
                st.write(f"Received: {mail.ReceivedTime}")
                st.write(mail.Body[:500] + '...')

                if st.button(f"Export This Email: {mail.Subject}", key=mail.EntryID):
                    export_email(mail, chain[0].ConversationTopic)
                    logger.info(f"Exported: {mail.Subject}")
                    st.success("Exported successfully!")

            if st.button(f"Export Entire Conversation", key=cid):
                export_conversation(chain, chain[0].ConversationTopic)
                logger.info(f"Exported Conversation: {chain[0].ConversationTopic}")
                st.success("Conversation Exported successfully!")


# config.py
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # The directory where config.py is located

EXPORT_PATH = os.path.join(BASE_DIR, 'EmailExports')

# New & Modern way
# from pathlib import Path

# BASE_DIR = Path(__file__).resolve().parent
# EXPORT_PATH = BASE_DIR / 'EmailExports'



# logger.py
import logging

def get_logger():
    logging.basicConfig(
        filename='explorer.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return logging.getLogger()


# email_client.py
import win32com.client

class OutlookClient:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)

    def search_emails(self, keyword):
        items = self.inbox.Items
        items.Sort("[ReceivedTime]", True)

        results = []
        for mail in items:
            if mail.Class == 43:
                if keyword.lower() in mail.Subject.lower() or keyword.lower() in mail.Body.lower():
                    results.append(mail)
        return results


# chain_handler.py
from collections import defaultdict


def group_by_conversation(emails):
    grouped = defaultdict(list)
    for mail in emails:
        grouped[mail.ConversationID].append(mail)
    return grouped


# exporter.py
import os
from datetime import datetime
from config import EXPORT_PATH
from utils import safe_filename


def export_email(mail, conversation_topic):
    # Safe Folder Name
    safe_folder_name = safe_filename(conversation_topic)
    folder_path = os.path.join(EXPORT_PATH, safe_folder_name)
    os.makedirs(folder_path, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    safe_file_name = f"{safe_filename(mail.Subject)}_{timestamp}.txt"
    file_path = os.path.join(folder_path, safe_file_name)

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(f"Subject: {mail.Subject}\n")
        f.write(f"From: {mail.SenderEmailAddress}\n")
        f.write(f"To: {mail.To}\n")
        f.write(f"Date: {mail.ReceivedTime}\n\n")
        f.write(mail.Body)

    # Export Attachments
    for att in mail.Attachments:
        att.SaveAsFile(os.path.join(folder_path, att.FileName))


def export_conversation(chain, conversation_topic):
    safe_folder_name = safe_filename(conversation_topic)
    folder_path = os.path.join(EXPORT_PATH, safe_folder_name)
    os.makedirs(folder_path, exist_ok=True)

    file_name = f"Conversation_{safe_filename(conversation_topic)}.txt"
    file_path = os.path.join(folder_path, file_name)

    sorted_chain = sorted(chain, key=lambda x: x.ReceivedTime)

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(f"=== Conversation Export: {conversation_topic} ===\n\n")

        for mail in sorted_chain:
            f.write("=" * 40 + "\n")
            f.write(f"Subject: {mail.Subject}\n")
            f.write(f"From: {mail.SenderEmailAddress}\n")
            f.write(f"To: {mail.To}\n")
            f.write(f"Date: {mail.ReceivedTime}\n\n")
            f.write(mail.Body)
            f.write("\n\n")

            for att in mail.Attachments:
                att.SaveAsFile(os.path.join(folder_path, att.FileName))



# utils
import re

def safe_filename(name, max_length=100):
    # Remove invalid characters
    name = re.sub(r'[\\/*?:"<>|]', "", name)

    # Trim length if too long
    if len(name) > max_length:
        name = name[:max_length]

    return name.strip()


# requirements.txt
pywin32
streamlit
pandas
