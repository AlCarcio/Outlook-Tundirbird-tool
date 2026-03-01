import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import threading
import os
import sys

# Append root directory to path to import converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from converter.msg_to_eml import MsgToEmlConverter


class MsgApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("MSG to EML Converter - Gilgatech Edition")
        self.geometry("800x600")

        # Set theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.is_converting = False

        # Mode selection: single file or folder
        self.mode_var = tk.StringVar(value="folder")

        self.lbl_mode = ctk.CTkLabel(self, text="Mode:")
        self.lbl_mode.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")

        self.mode_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.mode_frame.grid(row=0, column=1, padx=20, pady=(20, 5), sticky="w")

        self.rb_folder = ctk.CTkRadioButton(self.mode_frame, text="Folder (batch)", variable=self.mode_var, value="folder", command=self._on_mode_change)
        self.rb_folder.pack(side="left", padx=(0, 20))

        self.rb_file = ctk.CTkRadioButton(self.mode_frame, text="Single File", variable=self.mode_var, value="file", command=self._on_mode_change)
        self.rb_file.pack(side="left")

        # Input Selection
        self.lbl_input = ctk.CTkLabel(self, text="Select MSG Folder:")
        self.lbl_input.grid(row=1, column=0, padx=20, pady=10, sticky="w")

        self.entry_input = ctk.CTkEntry(self, textvariable=self.input_path)
        self.entry_input.grid(row=1, column=1, padx=20, pady=10, sticky="ew")

        self.btn_input = ctk.CTkButton(self, text="Browse", command=self.browse_input)
        self.btn_input.grid(row=1, column=2, padx=20, pady=10)

        # Output Directory Selection
        self.lbl_out = ctk.CTkLabel(self, text="Output Directory:")
        self.lbl_out.grid(row=2, column=0, padx=20, pady=10, sticky="w")

        self.entry_out = ctk.CTkEntry(self, textvariable=self.output_path)
        self.entry_out.grid(row=2, column=1, padx=20, pady=10, sticky="ew")

        self.btn_out = ctk.CTkButton(self, text="Browse", command=self.browse_output)
        self.btn_out.grid(row=2, column=2, padx=20, pady=10)

        # Convert Button
        self.btn_convert = ctk.CTkButton(self, text="Start Conversion", command=self.start_conversion, fg_color="green", hover_color="darkgreen")
        self.btn_convert.grid(row=3, column=0, columnspan=3, padx=20, pady=20)

        # Progress Label
        self.lbl_status = ctk.CTkLabel(self, text="Ready")
        self.lbl_status.grid(row=4, column=0, columnspan=3, padx=20, pady=(0, 10))

        # Log Window
        self.log_box = ctk.CTkTextbox(self, activate_scrollbars=True)
        self.log_box.grid(row=5, column=0, columnspan=3, padx=20, pady=20, sticky="nsew")
        self.log_box.configure(state="disabled")

    def _on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "folder":
            self.lbl_input.configure(text="Select MSG Folder:")
        else:
            self.lbl_input.configure(text="Select MSG File:")
        self.input_path.set("")

    def browse_input(self):
        mode = self.mode_var.get()
        if mode == "folder":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Outlook Message", "*.msg")])
        if path:
            self.input_path.set(path)

    def browse_output(self):
        dirname = filedialog.askdirectory()
        if dirname:
            self.output_path.set(dirname)

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def update_status(self, text):
        self.lbl_status.configure(text=text)

    def start_conversion(self):
        if self.is_converting:
            return

        input_p = self.input_path.get()
        out = self.output_path.get()

        if not input_p or not os.path.exists(input_p):
            self.log("[Error] Invalid input path.")
            return
        if not out or not os.path.exists(out):
            self.log("[Error] Invalid output directory.")
            return

        self.is_converting = True
        self.btn_convert.configure(state="disabled", text="Converting...")
        self.update_status("Converting...")
        self.log(f"Starting conversion: {input_p} -> {out}")

        threading.Thread(target=self.run_conversion_thread, args=(input_p, out), daemon=True).start()

    def run_conversion_thread(self, input_path, output_dir):
        try:
            converter = MsgToEmlConverter()
            converter.convert(input_path, output_dir, self.log_callback)
            self.log("[Success] Conversion completed successfully.")
            self.update_status("Conversion Complete")
        except Exception as e:
            self.log(f"[Error] An error occurred: {str(e)}")
            self.update_status("Error Occurred")
        finally:
            self.is_converting = False
            self.btn_convert.configure(state="normal", text="Start Conversion")

    def log_callback(self, message):
        self.after(0, self.log, message)


if __name__ == "__main__":
    app = MsgApp()
    app.mainloop()
