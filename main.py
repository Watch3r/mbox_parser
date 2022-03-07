# Based on: https://gist.github.com/benwattsjones/060ad83efd2b3afc8b229d41f9b246c4
# Fixed a few things that weren't working for content types and logic.

import base64
import hashlib
import mailbox
from bs4 import BeautifulSoup


class GmailMboxMessage():
    def __init__(self, email_data, get_content: bool = True, hash_content: bool = True):
        if not isinstance(email_data, mailbox.mboxMessage):
            raise TypeError('Variable must be type mailbox.mboxMessage')
        self.email_data = email_data
        self.hash_content = hash_content
        self.get_content = get_content

    def parse_email(self):
        return {'labels': self.email_data['X-Gmail-Labels'],
                'date': self.email_data['Date'],
                'from': self._clean_emails('From'),
                'reply-to': self._clean_emails('Reply-to'),
                'to': self._clean_emails('To'),
                'delivered-to': self._clean_emails('Delivered-To'),
                'cc': self._clean_emails('Cc'),
                'bcc': self._clean_emails('Bcc'),
                'subject': self.email_data['Subject'],
                'payload': self.read_email_payload() if self.get_content else False,
                'message-id': self.email_data['Message-ID']}

    def read_email_payload(self):
        email_payload = self.email_data.get_payload()
        if self.email_data.is_multipart():
            email_messages = list(self._get_email_messages(email_payload))
        else:
            email_messages = [email_payload]
        return [self._read_email_text(msg) for msg in email_messages]

    def _get_email_messages(self, email_payload):
        for msg in email_payload:
            if isinstance(msg, (list,tuple)):
                for submsg in self._get_email_messages(msg):
                    yield submsg
            elif msg.is_multipart():
                for submsg in self._get_email_messages(msg.get_payload()):
                    yield submsg
            else:
                yield msg

    def _read_email_text(self, msg):
        # Find content type from msg
        content_type = 'NA' if isinstance(msg, str) else msg.get_content_type()
        # Find encoding used in msg
        encoding = 'NA' if isinstance(msg, str) else msg.get('Content-Transfer-Encoding', 'NA')

        # Parse messaging based on content type
        # Note: If you're getting message type as None, you should probably do some debugging here.
        if content_type == 'text/plain' and encoding == 'base64':
            msg_text = msg.get_payload()
        elif content_type in ['text/html', 'application/octet-stream'] and encoding == 'base64':
            msg_text = self._get_html_text(msg.get_payload())
        elif content_type == 'NA':
            msg_text = self._get_html_text(msg)
        else:
            msg_text = None

        try:
            file_name = msg.get_filename()
        except AttributeError:
            file_name = None

        msg_text = msg_text.replace("\n", "") if isinstance(msg_text, str) else msg_text

        if self.hash_content:
            content_hash = hash_string(msg_text) if type(msg_text) is str else None
            return {'content_type': content_type, 'encoding': encoding, 'file_name': file_name, 'content': msg_text,
                    'content_hash': content_hash}

        return {'content_type': content_type, 'encoding': encoding, 'file_name': file_name, 'content': msg_text,
                'content_hash': False}

    def _get_html_text(self, html):
        try:
            return BeautifulSoup(html, 'lxml').body.get_text(' ', strip=True)
        except AttributeError:  # message contents empty
            return None

    def _clean_emails(self, field: str):
        if self.email_data[field] is None:
            return []

        field = self.email_data[field].lower()

        # If any strings pop up in your emails you don't want, add it to the strip_strings. Removes some extras for what is required.
        # https://stackoverflow.com/questions/2049502/what-characters-are-allowed-in-an-email-address
        strip_strings = ["\n", "\t", "<", ">", "\"", "\'", ","]
        for strip in strip_strings:
            field = field.replace(strip, " ")

        emails = [e for e in field.split(" ") if "@" in e and e.count(".") >= 1]

        return sorted(list(set(emails)))


def write_b64_file(file_name: str, b64_text: str):
    with open(file_name, "wb") as fh:
        fh.write(base64.decodebytes(b64_text.encode('ascii')))


def hash_string(text: str):
    # MD5 Is current hashing algorithm used.
    return hashlib.md5(text.encode('utf-8')).hexdigest()


def main():
    mbox_obj = mailbox.mbox(r'path\to\file.mbox')
    num_emails = len(mbox_obj)

    for index, email_obj in enumerate(mbox_obj, start=1):
        email_data = GmailMboxMessage(email_obj, get_content=True, hash_content=True).parse_email()
        print(f"{index}/{num_emails}", email_data)


if __name__ == '__main__':
    main()
