import customtkinter as ctk
from tkinter import filedialog
import os
import json
import threading
import webbrowser
from scraper import run_scraper
from transaction_scraper import run_transaction_scraper

CONFIG_FILE = "config.json"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("GW2 Trading Scraper")
        self.geometry("700x600")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.output_dir = os.path.abspath('.')
        self.api_key = ""
        self.load_config()

        # Scraper Frame
        self.scraper_frame = ctk.CTkFrame(self)
        self.scraper_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.scraper_frame.grid_columnconfigure(0, weight=1) # Center content
        self.scraper_label = ctk.CTkLabel(self.scraper_frame, text="Item Flip Scraper", font=ctk.CTkFont(size=16, weight="bold"))
        self.scraper_label.grid(row=0, column=0, padx=10, pady=10)

        # Arguments Frame for Scraper
        self.scraper_args_frame = ctk.CTkFrame(self.scraper_frame)
        self.scraper_args_frame.grid(row=1, column=0, pady=5)
        self.historical_check = ctk.CTkCheckBox(self.scraper_args_frame, text="Fetch Historical Data (slower)")
        self.historical_check.grid(row=0, column=0, padx=10, pady=5)
        self.days_label = ctk.CTkLabel(self.scraper_args_frame, text="Days of history:")
        self.days_label.grid(row=0, column=1, padx=10, pady=5)
        self.days_entry = ctk.CTkEntry(self.scraper_args_frame, width=50)
        self.days_entry.grid(row=0, column=2, padx=10, pady=5)
        self.days_entry.insert(0, "7")
        self.pages_label = ctk.CTkLabel(self.scraper_args_frame, text="Pages to scrape (0 for all):")
        self.pages_label.grid(row=0, column=3, padx=10, pady=5)
        self.pages_entry = ctk.CTkEntry(self.scraper_args_frame, width=50)
        self.pages_entry.grid(row=0, column=4, padx=10, pady=5)
        self.pages_entry.insert(0, "0")

        self.run_scraper_button = ctk.CTkButton(self.scraper_frame, text="Run Scraper", command=self.start_scraper_thread)
        self.run_scraper_button.grid(row=2, column=0, padx=10, pady=10)

        # Transaction Scraper Frame
        self.transaction_frame = ctk.CTkFrame(self)
        self.transaction_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.transaction_frame.grid_columnconfigure(0, weight=1) # Center content
        self.transaction_label = ctk.CTkLabel(self.transaction_frame, text="Profit & Loss Report", font=ctk.CTkFont(size=16, weight="bold"))
        self.transaction_label.grid(row=0, column=0, padx=10, pady=10)

        # Arguments Frame for Transaction Scraper
        self.transaction_args_frame = ctk.CTkFrame(self.transaction_frame)
        self.transaction_args_frame.grid(row=1, column=0, pady=5)
        self.api_key_label = ctk.CTkLabel(self.transaction_args_frame, text="GW2 API Key:")
        self.api_key_label.grid(row=0, column=0, padx=10, pady=5)
        self.api_key_entry = ctk.CTkEntry(self.transaction_args_frame, placeholder_text="Enter your API key here", width=350)
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=5)
        self.api_key_entry.insert(0, self.api_key)
        self.trans_days_label = ctk.CTkLabel(self.transaction_args_frame, text="Days of history:")
        self.trans_days_label.grid(row=1, column=0, padx=10, pady=5)
        self.trans_days_entry = ctk.CTkEntry(self.transaction_args_frame, width=50)
        self.trans_days_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        self.trans_days_entry.insert(0, "30")

        self.run_transaction_button = ctk.CTkButton(self.transaction_frame, text="Run Profit Report", command=self.start_transaction_thread)
        self.run_transaction_button.grid(row=2, column=0, padx=10, pady=10)

        # Settings Frame
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.settings_frame.grid_columnconfigure(1, weight=1)
        self.settings_label = ctk.CTkLabel(self.settings_frame, text="Settings", font=ctk.CTkFont(size=16, weight="bold"))
        self.settings_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        self.output_dir_button = ctk.CTkButton(self.settings_frame, text="Choose Output Folder", command=self.choose_output_dir)
        self.output_dir_button.grid(row=1, column=0, padx=10, pady=10)
        self.output_dir_label = ctk.CTkLabel(self.settings_frame, text=f"Output Folder: {self.output_dir}", anchor="w")
        self.output_dir_label.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Log Frame
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.log_textbox = ctk.CTkTextbox(self.log_frame, state="disabled", wrap="word")
        self.log_textbox.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def log(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"{message}\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

    def safe_log(self, message):
        self.after(0, self.log, message)

    def load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.api_key = config.get("api_key", "")
                    self.output_dir = config.get("output_dir", os.path.abspath('.'))
        except (IOError, json.JSONDecodeError) as e:
            self.log(f"Could not load config: {e}")

    def save_config(self):
        try:
            config = {
                "api_key": self.api_key_entry.get(),
                "output_dir": self.output_dir
            }
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            self.log(f"Could not save config: {e}")

    def choose_output_dir(self):
        dir_path = filedialog.askdirectory(initialdir=self.output_dir)
        if dir_path:
            self.output_dir = os.path.abspath(dir_path)
            self.output_dir_label.configure(text=f"Output Folder: {self.output_dir}")
            self.log(f"Output directory set to: {self.output_dir}")

    def set_buttons_state(self, state):
        self.run_scraper_button.configure(state=state)
        self.run_transaction_button.configure(state=state)
        self.output_dir_button.configure(state=state)

    def start_scraper_thread(self):
        self.set_buttons_state("disabled")
        self.log("--- Starting Item Flip Scraper ---")
        historical = self.historical_check.get()
        days = 7
        try:
            days = int(self.days_entry.get())
        except ValueError:
            self.log("Invalid input for days. Using default of 7.")
            days = 7

        pages = 0
        try:
            pages = int(self.pages_entry.get())
        except ValueError:
            self.log("Invalid input for pages. Using default of 0 (all).")
            pages = 0

        thread = threading.Thread(target=run_scraper, args=(historical, self.output_dir, days, pages, self.safe_log))
        thread.daemon = True
        thread.start()
        self.monitor_thread(thread, task_type="scraper", output_dir=self.output_dir)

    def start_transaction_thread(self):
        self.set_buttons_state("disabled")
        self.log("--- Starting Profit & Loss Report ---")
        api_key = self.api_key_entry.get()
        if not api_key:
            self.log("Error: API Key is required.")
            self.set_buttons_state("normal")
            return

        days = 30
        try:
            days = int(self.trans_days_entry.get())
        except ValueError:
            self.log("Invalid input for days. Using default of 30.")
            days = 30

        thread = threading.Thread(target=run_transaction_scraper, args=(api_key, self.output_dir, self.safe_log, days))
        thread.daemon = True
        thread.start()
        self.monitor_thread(thread, task_type="transaction_scraper", output_dir=self.output_dir)

    def show_dashboard(self, output_dir):
        report_path = os.path.join(output_dir, "interactive_report.html")
        if not os.path.exists(report_path):
            self.log(f"Error: Could not find report file at {report_path}")
            return

        try:
            # Create a file:// URL
            report_url = f"file:///{os.path.abspath(report_path)}"
            webbrowser.open(report_url)
            self.log(f"Opening interactive report in your default web browser.")
        except Exception as e:
            self.log(f"Error opening web browser: {e}")

    def monitor_thread(self, thread, task_type=None, output_dir=None):
        if thread.is_alive():
            self.after(100, lambda: self.monitor_thread(thread, task_type, output_dir))
        else:
            self.log("--- Task Finished ---")
            self.set_buttons_state("normal")
            if task_type == "transaction_scraper":
                self.show_dashboard(output_dir)

    def on_closing(self):
        self.save_config()
        self.destroy()

if __name__ == "__main__":
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
