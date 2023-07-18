# This is a modification of the following code:
# https://pypi.org/project/msgtopdf/
# MIT License

import os
import re
import subprocess
import sys
from pathlib import Path, PurePath

import win32com.client


class MsgToPdf:

    def __init__(self, msg_path, pdf_path, tmp_folder):
        self.wkhtmltopdf = "wkhtmltopdf"
        if self.__check_path_exist() is False:
            sys.exit(1)
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.msg_path = msg_path
        self.pdf_path = pdf_path
        self.tmp_folder = tmp_folder
        self.html_path = Path(tmp_folder, Path(msg_path).stem + ".html")
        self.msg = outlook.OpenSharedItem(self.msg_path)

    def extract_email_attachments(self):
        count_attachments = self.msg.Attachments.Count
        if count_attachments > 0:
            for item in range(count_attachments):
                attachment_filename = self.msg.Attachments.Item(item + 1).Filename
                self.msg.Attachments.Item(item + 1).SaveAsFile(PurePath(self.tmp_folder, attachment_filename))

    def msg_to_pdf(self):
        email_body = self.__add_header_information() + self.raw_email_body()
        clean_email_body = self.replace_CID(email_body)
        self.extract_email_attachments()
        with open(self.html_path, "w", encoding="utf-8") as f:
            f.write(clean_email_body)
        subprocess.run([
            str(self.wkhtmltopdf),
            "--enable-local-file-access",
            "--log-level",
            "warn",
            "--encoding",
            "utf-8",
            "--footer-font-size",
            "6",
            "--footer-line",
            "--footer-center",
            "[page] / [topage]",
            str(self.html_path),
            str(self.pdf_path),
        ])

    def raw_email_body(self):
        if self.msg.BodyFormat == 2:
            body = self.msg.HTMLBody
            self.email_format = "html"
        elif self.msg.BodyFormat == 3:
            body = self.msg.RTFBody
            self.email_format = "html"
        else:
            body = self.msg.Body
            self.email_format = "txt"
        self.raw_body = body
        return self.raw_body

    def replace_CID(self, body):
        self.image_files = []
        # search for cid:(capture_group)@* upto "
        p = re.compile(r"cid:([^\"@]*)[^\"]*")
        r = p.sub(self.__return_image_reference, body)
        return r

    def __add_header_information(self):
        html_str = """
        <head>
        <meta charset="UTF-8">
        <base href="{base_href}">
        <p style="font-family: Arial;font-size: 11.0pt">
        </head>
        <strong>From:</strong>               {sender}</br>
        <strong>Sent:</strong>               {sent}</br>
        <strong>To:</strong>                 {to}</br>
        <strong>Cc:</strong>                 {cc}</br>
        <strong>Subject:</strong>            {subject}</p>
        <hr>
        """
        formatted_html = html_str.format(
            base_href="file:///" + str(self.tmp_folder) + "\\",
            sender=self.msg.SenderName,
            sent=self.msg.SentOn,
            to=self.msg.To,
            cc=self.msg.CC,
            subject=self.msg.Subject,
            attachments=self.msg.Attachments
        )
        return formatted_html

    def __check_path_exist(self):
        path = os.getenv("PATH")
        required_paths = [self.wkhtmltopdf]
        for p in required_paths:
            if p not in path:
                return False
        return True

    def __return_image_reference(self, match):
        value = str(match.groups()[0])
        if value not in self.image_files:
            self.image_files.append(value)
        return value


def msg_to_pdf(msg_path, pdf_path, tmp_folder):
    obj = MsgToPdf(msg_path, pdf_path, tmp_folder)
    obj.msg_to_pdf()
    return True
