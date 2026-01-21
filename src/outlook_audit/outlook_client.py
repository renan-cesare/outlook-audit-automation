import time
from dataclasses import dataclass
from typing import Optional, Tuple

import win32com.client as win32


@dataclass
class SentIds:
    conversation_id: str
    internet_message_id: str
    entry_id: str


class OutlookClient:
    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.sent_folder = self.namespace.GetDefaultFolder(5)   # Sent Items
        self.inbox_folder = self.namespace.GetDefaultFolder(6)  # Inbox

    def create_mail(self):
        return self.outlook.CreateItem(0)

    def send_mail(
        self,
        to: str,
        cc: str,
        subject: str,
        body: str,
        display_only: bool = False,
        is_html: bool = False,
    ) -> None:
        mail = self.create_mail()
        mail.To = to or ""
        mail.CC = cc or ""
        mail.Subject = subject or ""

        if is_html:
            mail.HTMLBody = body
        else:
            mail.Body = body

        # salva antes (ajuda a estabilizar IDs e reduzir casos “fantasma”)
        try:
            mail.Save()
        except Exception:
            pass

        if display_only:
            mail.Display()
        else:
            mail.Send()

    def find_sent_ids_by_subject_and_token(
        self,
        subject: str,
        token: str,
        delay_seconds: int = 3,
        max_items: int = 300,
    ) -> SentIds:
        if delay_seconds > 0:
            time.sleep(delay_seconds)

        items = self.sent_folder.Items
        try:
            items.Sort("[SentOn]", True)  # desc
        except Exception:
            pass

        conversation_id = ""
        internet_id = ""
        entry_id = ""

        count = 0
        for item in items:
            count += 1
            if count > max_items:
                break

            try:
                if getattr(item, "Subject", "") != subject:
                    continue

                body_text = getattr(item, "Body", "") or ""
                html_text = getattr(item, "HTMLBody", "") or ""

                if (token in body_text) or (token in html_text):
                    conversation_id = getattr(item, "ConversationID", "") or ""
                    internet_id = getattr(item, "InternetMessageID", "") or ""
                    entry_id = getattr(item, "EntryID", "") or ""
                    break
            except Exception:
                continue

        return SentIds(conversation_id=conversation_id, internet_message_id=internet_id, entry_id=entry_id)

    def get_item_from_id(self, entry_id: str):
        return self.namespace.GetItemFromID(entry_id)

    def scan_inbox_for_reply_by_conversation_id(
        self,
        conversation_id: str,
        max_items: int = 5000,
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        items = self.inbox_folder.Items
        try:
            items.Sort("[ReceivedTime]", True)  # desc
        except Exception:
            pass

        count = 0
        for item in items:
            count += 1
            if count > max_items:
                break
            try:
                if getattr(item, "Class", None) != 43:  # MailItem
                    continue
                if getattr(item, "ConversationID", "") == conversation_id:
                    received = getattr(item, "ReceivedTime", None)
                    body = getattr(item, "Body", "") or ""
                    received_iso = str(received) if received else None
                    return True, received_iso, body
            except Exception:
                continue

        return False, None, None
