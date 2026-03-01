import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import threading
import os
import sys

# Append root directory to path to import converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from converter.pst_to_mbox import PstConverter

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PST to MBOX Converter - Gilgatech Edition")
        self.geometry("800x600")
        
        # Set theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Variables
        self.pst_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.is_converting = False

        # Input PST Selection
        self.lbl_pst = ctk.CTkLabel(self, text="Select PST File:")
        self.lbl_pst.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.entry_pst = ctk.CTkEntry(self, textvariable=self.pst_path)
        self.entry_pst.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ew")
        
        self.btn_pst = ctk.CTkButton(self, text="Browse", command=self.browse_pst)
        self.btn_pst.grid(row=0, column=2, padx=20, pady=(20, 10))

        # Output Directory Selection
        self.lbl_out = ctk.CTkLabel(self, text="Output Directory:")
        self.lbl_out.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.entry_out = ctk.CTkEntry(self, textvariable=self.output_path)
        self.entry_out.grid(row=1, column=1, padx=20, pady=10, sticky="ew")
        
        self.btn_out = ctk.CTkButton(self, text="Browse", command=self.browse_output)
        self.btn_out.grid(row=1, column=2, padx=20, pady=10)

        # Convert Button
        self.btn_convert = ctk.CTkButton(self, text="Start Conversion", command=self.start_conversion, fg_color="green", hover_color="darkgreen")
        self.btn_convert.grid(row=2, column=0, columnspan=3, padx=20, pady=20)

        # Progress Label
        self.lbl_status = ctk.CTkLabel(self, text="Ready")
        self.lbl_status.grid(row=3, column=0, columnspan=3, padx=20, pady=(0, 10))

        # Log Window
        self.log_box = ctk.CTkTextbox(self, activate_scrollbars=True)
        self.log_box.grid(row=4, column=0, columnspan=3, padx=20, pady=20, sticky="nsew")
        self.log_box.configure(state="disabled")

    def browse_pst(self):
        filename = filedialog.askopenfilename(filetypes=[("Outlook Data File", "*.pst")])
        if filename:
            self.pst_path.set(filename)

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

        pst = self.pst_path.get()
        out = self.output_path.get()

        if not pst or not os.path.exists(pst):
            self.log("[Error] Invalid PST file path.")
            return
        if not out or not os.path.exists(out):
            self.log("[Error] Invalid output directory.")
            return

        self.is_converting = True
        self.btn_convert.configure(state="disabled", text="Converting...")
        self.log(f"Starting conversion: {pst} -> {out}")

        # Run in thread to keep UI responsive
        threading.Thread(target=self.run_conversion_thread, args=(pst, out), daemon=True).start()

    def run_conversion_thread(self, pst_path, output_dir):
        try:
            converter = PstConverter()
            converter.convert(pst_path, output_dir, self.log_callback)
            self.log("[Success] Conversion completed successfully.")
            self.update_status("Conversion Complete")
        except Exception as e:
            self.log(f"[Error] An error occurred: {str(e)}")
            self.update_status("Error Occurred")
        finally:
            self.is_converting = False
            self.btn_convert.configure(state="normal", text="Start Conversion")

    def log_callback(self, message):
        # Schedule UI update on main thread (simple approach for Tkinter, usually safe for simple inserts)
        self.after(0, self.log, message)

if __name__ == "__main__":
    app = App()
    app.mainloop()
