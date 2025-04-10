# Streamlit Outlook Email Explorer - Project Structure

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
EXPORT_PATH = r"C:\\EmailExports"


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
from config import EXPORT_PATH
from utils import safe_filename  

#file_name = safe_filename(mail.Subject)
#file_path = os.path.join(folder_name, f"{file_name}.txt")

def export_email(mail, conversation_topic):
    folder_name = f"{EXPORT_PATH}\\{conversation_topic.replace(' ','_')}"
    os.makedirs(folder_name, exist_ok=True)

    file_path = f"{folder_name}\\{mail.Subject[:50].replace(' ','_')}.txt"

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(f"Subject: {mail.Subject}\n")
        f.write(f"From: {mail.SenderEmailAddress}\n")
        f.write(f"Date: {mail.ReceivedTime}\n\n")
        f.write(mail.Body)

    for att in mail.Attachments:
        att.SaveAsFile(os.path.join(folder_name, att.FileName))


def export_conversation(chain, conversation_topic):
    folder_name = f"{EXPORT_PATH}\\{conversation_topic.replace(' ','_')}"
    os.makedirs(folder_name, exist_ok=True)

    file_path = f"{folder_name}\\Conversation_{conversation_topic[:50].replace(' ','_')}.txt"

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
                att.SaveAsFile(os.path.join(folder_name, att.FileName))

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
