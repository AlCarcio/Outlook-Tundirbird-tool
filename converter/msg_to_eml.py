import os
import email
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
            eml_message = email.message.EmailMessage()

            # Set headers
            if msg.subject:
                eml_message["Subject"] = msg.subject
            if msg.sender:
                eml_message["From"] = msg.sender
            if msg.to:
                eml_message["To"] = msg.to
            if msg.cc:
                eml_message["Cc"] = msg.cc
            if msg.date:
                eml_message["Date"] = str(msg.date)
            if msg.messageId:
                eml_message["Message-ID"] = msg.messageId

            # Set body
            if msg.htmlBody:
                eml_message.set_content(msg.body or "", subtype="plain")
                eml_message.add_alternative(msg.htmlBody, subtype="html")
            elif msg.body:
                eml_message.set_content(msg.body, subtype="plain")

            # Handle attachments
            for attachment in msg.attachments:
                try:
                    att_name = attachment.longFilename or attachment.shortFilename or "attachment"
                    att_data = attachment.data
                    if att_data:
                        maintype, _, subtype = (attachment.mimetype or "application/octet-stream").partition("/")
                        if not subtype:
                            maintype = "application"
                            subtype = "octet-stream"
                        eml_message.add_attachment(
                            att_data,
                            maintype=maintype,
                            subtype=subtype,
                            filename=att_name
                        )
                except Exception as e:
                    self._log(f"    [Warning] Could not attach '{att_name}': {e}")

            # Write EML file
            with open(eml_path, "w", encoding="utf-8") as f:
                f.write(eml_message.as_string())

            self._log(f"    -> Saved: {eml_filename}")

        finally:
            msg.close()

    def _sanitize_filename(self, name):
        return "".join([c for c in name if c.isalpha() or c.isdigit() or c in (' ', '-', '_', '.')]).strip()

    def _log(self, text):
        if self.callback:
            self.callback(text)
        print(text)
