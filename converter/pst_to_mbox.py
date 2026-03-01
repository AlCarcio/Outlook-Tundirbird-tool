import os
import io
import time
# Aspose.Email Imports
import aspose.email as ae
from aspose.email.storage.pst import PersonalStorage
from aspose.email.mapi import MailConversionOptions

class PstConverter:
    def convert(self, pst_path, output_dir, progress_callback=None):
        """
        Converts a PST file to MBOX files in the output directory.
        """
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        self.callback = progress_callback
        
        try:
            self._log(f"Loading PST file: {pst_path}")
            self._log("Note: Using Aspose.Email Evaluation Version (Limitations may apply)")
            
            # Load PST
            with PersonalStorage.from_file(pst_path) as pst:
                # Recursively process folders
                self._process_folder(pst.root_folder, pst, output_dir)
                
        except Exception as e:
            self._log(f"Critical Error: {e}")
            raise

    def _process_folder(self, folder_info, pst, output_dir_base):
        """
        Recursively process folders and save messages.
        """
        # Get folder name
        display_name = folder_info.display_name
        
        # Skip root container if empty name
        if not display_name:
            display_name = "Root"

        safe_name = self._sanitize_filename(display_name)
        
        # Log every folder found to debug structure
        item_count = folder_info.content_count
        self._log(f"DEBUG: Found folder '{display_name}' with {item_count} items")

        # Only process if it has content
        if item_count > 0:
            mbox_filename = f"{safe_name}.mbox"
            mbox_path = os.path.join(output_dir_base, mbox_filename)
            
            message_count = 0
            
            try:
                # Open MBOX file for appending
                with open(mbox_path, "ab") as mbox_file:
                    for msg_info in folder_info.get_contents():
                        try:
                            # Extract message as MapiMessage
                            mapi_msg = pst.extract_message(msg_info)
                            
                            # Convert to MailMessage for EML export
                            # We need standard MailMessage to use standard SaveOptions
                            mail_msg = mapi_msg.to_mail_message(MailConversionOptions())

                            # Write "From " separator
                            date_str = time.ctime() 
                            try:
                                if mail_msg.date:
                                    date_str = mail_msg.date.strftime("%a %b %d %H:%M:%S %Y")
                            except:
                                pass
                                
                            separator = f"From - {date_str}\r\n".encode("utf-8")
                            mbox_file.write(separator)
                            
                            # Save to stream using default EML format
                            stream = io.BytesIO()
                            mail_msg.save(stream, ae.SaveOptions.default_eml)
                            eml_bytes = stream.getvalue()
                            
                            mbox_file.write(eml_bytes)
                            mbox_file.write(b"\r\n\r\n") 
                            
                            message_count += 1
                            
                        except Exception as e:
                           self._log(f"  [Warning] Failed to extract a message: {e}")
            except Exception as e:
                self._log(f"  [Error] Could not write to file {mbox_path}: {e}")
                       
            if message_count > 0:
                self._log(f"  -> Saved {message_count} messages to {mbox_filename}")
            else:
                self._log(f"  -> No readable messages extracted from {display_name}")

        # Recurse
        if folder_info.has_sub_folders:
            for sub_folder in folder_info.get_sub_folders():
                self._process_folder(sub_folder, pst, output_dir_base)

    def _sanitize_filename(self, name):
        return "".join([c for c in name if c.isalpha() or c.isdigit() or c in (' ', '-', '_', '.')]).strip()

    def _log(self, text):
        if self.callback:
            self.callback(text)
        print(text)
