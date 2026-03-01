import os
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import extract_msg


class MsgToEmlConverter:
    def __init__(self):
        self.callback = None

    def convert(self, input_path, output_dir, progress_callback=None):
        """
        Converts MSG files to EML format.
        input_path can be a single .msg file or a directory containing .msg files.
        """
        self.callback = progress_callback

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Collect all .msg files
        msg_files = []
        if os.path.isfile(input_path):
            if input_path.lower().endswith(".msg"):
                msg_files.append(input_path)
            else:
                self._log(f"[Error] File is not a .msg file: {input_path}")
                return
        elif os.path.isdir(input_path):
            for root, dirs, files in os.walk(input_path):
                for f in files:
                    if f.lower().endswith(".msg"):
                        msg_files.append(os.path.join(root, f))
        else:
            self._log(f"[Error] Path not found: {input_path}")
            return

        if not msg_files:
            self._log("[Warning] No .msg files found.")
            return

        self._log(f"Found {len(msg_files)} MSG file(s) to convert.")
        converted = 0
        errors = 0

        for msg_path in msg_files:
            try:
                self._convert_single(msg_path, output_dir)
                converted += 1
            except Exception as e:
                self._log(f"  [Error] Failed to convert {os.path.basename(msg_path)}: {e}")
                errors += 1

        self._log(f"\nConversion complete: {converted} converted, {errors} errors.")

    def _convert_single(self, msg_path, output_dir):
        """Convert a single MSG file to EML."""
        basename = os.path.splitext(os.path.basename(msg_path))[0]
        eml_filename = self._sanitize_filename(basename) + ".eml"
        eml_path = os.path.join(output_dir, eml_filename)

        # Handle duplicate filenames
        counter = 1
        while os.path.exists(eml_path):
            eml_filename = f"{self._sanitize_filename(basename)}_{counter}.eml"
            eml_path = os.path.join(output_dir, eml_filename)
            counter += 1

        self._log(f"  Converting: {os.path.basename(msg_path)}")

        # Open MSG and build EML
        msg = extract_msg.Message(msg_path)
        try:
            has_html = bool(msg.htmlBody)
            has_body = bool(msg.body)
            has_attachments = len(msg.attachments) > 0

            # Decide structure
            if has_attachments or (has_html and has_body):
                eml = MIMEMultipart("mixed")
            elif has_html:
                eml = MIMEMultipart("alternative")
            else:
                eml = MIMEMultipart()

            # Set headers
            if msg.subject:
                eml["Subject"] = msg.subject
            if msg.sender:
                eml["From"] = msg.sender
            if msg.to:
                eml["To"] = msg.to
            if msg.cc:
                eml["Cc"] = msg.cc
            if msg.date:
                eml["Date"] = str(msg.date)
            if msg.messageId:
                eml["Message-ID"] = msg.messageId

            # Add body
            if has_html and has_body:
                alt_part = MIMEMultipart("alternative")
                alt_part.attach(MIMEText(msg.body, "plain", "utf-8"))
                alt_part.attach(MIMEText(msg.htmlBody, "html", "utf-8"))
                eml.attach(alt_part)
            elif has_html:
                eml.attach(MIMEText(msg.htmlBody, "html", "utf-8"))
            elif has_body:
                eml.attach(MIMEText(msg.body, "plain", "utf-8"))

            # Handle attachments
            for attachment in msg.attachments:
                try:
                    att_name = attachment.longFilename or attachment.shortFilename or "attachment"
                    att_data = attachment.data
                    if att_data:
                        # Guess MIME type from filename
                        mime_type, _ = mimetypes.guess_type(att_name)
                        if not mime_type:
                            mime_type = "application/octet-stream"
                        maintype, subtype = mime_type.split("/", 1)

                        part = MIMEBase(maintype, subtype)
                        part.set_payload(att_data)
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", "attachment", filename=att_name)
                        eml.attach(part)
                except Exception as e:
                    self._log(f"    [Warning] Could not attach file: {e}")

            # Write EML file
            with open(eml_path, "wb") as f:
                f.write(eml.as_bytes())

            self._log(f"    -> Saved: {eml_filename}")

        finally:
            msg.close()

    def _sanitize_filename(self, name):
        return "".join([c for c in name if c.isalpha() or c.isdigit() or c in (' ', '-', '_', '.')]).strip()

    def _log(self, text):
        if self.callback:
            self.callback(text)
        print(text)
