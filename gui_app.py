import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import traceback
import threading
import configparser
import os
import main
from storage_gspread import SpreadsheetDocu


CONFIG_FILE = "settings.ini"


def load_config_file(config_file: str = CONFIG_FILE) -> configparser.ConfigParser:
    """
    Loads a configuration file. Creates it with default values, if it does not exist.

    Parameters:
        config_file (str): The path to the configuration file.

    Returns:
        configparser.ConfigParser: The loaded configuration.
    """
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
    else:
        config["General"] = {
            "installed": "0",
            "autostart": "0",
            "window_width": "500",
            "window_height": "400"
        }
        with open(config_file, "w") as f:
            config.write(f)
    return config


class App(tk.Tk):
    """
    Main application window.

    Attributes:
        config (configparser.ConfigParser): The loaded / created configuration.
        is_running (bool): Whether a process is running.
        autostart_val (tk.IntVar): The value of the autostart checkbox.
        check_autostart (tk.Checkbutton): Checkbox to run documentation updater at App-start.
        start_button (ttk.Button): The start documentation updater button.
        log_area (scrolledtext.ScrolledText): The log text area.
        install_button (ttk.Button): The start installer button.
    """
    def __init__(self):
        """
        Initializes the main application window. Loads the configuration file and sets up the window's properties and program state.
        """
        super().__init__()
        self.config = load_config_file(CONFIG_FILE)

        self.title("external drives Documentation Updater")
        self.geometry(
            str(self.config.getint('General', 'window_width'))
            + "x"
            + str(self.config.getint('General', 'window_height'))
        )

        # Scrolled text area for logs
        self.log_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=10, state="disabled")
        self.log_area.pack(fill="both", expand=False, padx=10, pady=10)

        # Installer
        if not self.config.getint("General", "installed"):
            self._log('To access the Google Spreadsheet service_account.json is needed.\n\n'
                      'To move service_account.json to its correct location, '
                      f'put it in this folder: \n{os.getcwd()}\nand click "Install Program"')
            self.install_button = ttk.Button(self, text="Install Program", command=self._installer)
            self.install_button.pack(pady=10)
            return

        self.is_running = False

        # checkbox for autostart
        self.autostart_val = tk.IntVar(value=int(self.config.getint("General", "autostart")))
        self.check_autostart = tk.Checkbutton(
            self,
            text="automatically update documentation at App-Start",
            variable=self.autostart_val,
            command=self._autostart_checkbox_change
        )
        self.check_autostart.pack()

        # button to manually start process
        self.start_button = ttk.Button(self, text="Scan Drives & Update Documentation", command=self._start_docu_update_thread)
        self.start_button.pack(pady=10)

        # logic
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        if self.autostart_val.get():
            self._start_docu_update_thread()

    def _autostart_checkbox_change(self):
        self.config["General"]["autostart"] = str(self.autostart_val.get())
        self._update_config()

    def _update_config(self):
        with open(CONFIG_FILE, "w") as f:
            self.config.write(f)

    def _on_closing(self):
        """Handles window closing event"""
        if self.is_running:
            message = (
                "Quitting while the documentation is updating may result in an empty documentation.\n\n"
                "(You can restore it from the Spreadsheet history.)\n\n"
                "Are you sure you want to quit?"
            )
            if not messagebox.askyesno(title="Quit?", message=message):
                return
        self.destroy()

    def _log(self, message):
        """
        Appends a message to the log window.

        Parameters:
            message (str): Text to append.
        """
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, f"{message}\n")
        self.log_area.configure(state="disabled")
        self.log_area.see(tk.END)  # auto-scroll to bottom
        self.update_idletasks()

    @staticmethod
    def _run_formatter(spread_utils):
        from storage_gspread import SpreadsheetFormatter

        spread_formatter = SpreadsheetFormatter(spread_utils)
        spread_formatter.format_drives_column()

    def _docu_update_worker(self):
        """Runs the documentation update process, logs the progress and cleans up prevention of running multiple processes"""
        try:
            spread_utils = SpreadsheetDocu(main.SPREADSHEET_NAME)

            self._log("Scanning drives...")
            blacklist_drives, blacklist_directories = spread_utils.load_blacklist()
            drives_df = main.scan_valid_drives_to_df(blacklist_drives, blacklist_directories)

            if drives_df.empty:
                self._log("No valid drives are found (they may be blacklisted)")
                self._log("Documentation is not updated.")
            else:
                self._log("Updating Spreadsheet...")
                spread_utils.update_online_spreadsheet(drives_df)

                self._log("Formatting Spreadsheet...")
                self._run_formatter(spread_utils)

                self._log("🎉 Documentation successfully updated!")

        except Exception as e:
            error_msg = f"❌ ERROR: {e}\n{traceback.format_exc()}"
            self._log(error_msg)

        finally:
            self.is_running = False
            self.start_button.config(state="normal")

    def _start_docu_update_thread(self):
        """Starts docu_update_worker() and prevents running multiple processes"""
        if self.is_running:
            self._log("⚠ A process is already running.")
            return

        self.is_running = True
        self.start_button.config(state="disabled")
        self._log("🚀 Starting process...")
        threading.Thread(target=self._docu_update_worker, daemon=True).start()

    def _installer(self):
        """Checks if service_account.json is in its correct location, or moves it there from working directory. Logs progress and errors."""
        import shutil
        from pathlib import Path

        self._log("\n\n")
        try:
            if os.name == "nt":
                home_config_dir = Path.home() / "AppData" / "Roaming" / "gspread"
            else:
                home_config_dir = Path.home() / ".config" / "gspread"
            service_account_correct_location = home_config_dir / "service_account.json"
            service_account_local = Path("service_account.json")

            if service_account_correct_location.exists():
                self._log(f"service_account.json is already in the correct location: {str(service_account_correct_location)}")
            elif not service_account_local.exists():
                self._log("service_account.json not found.")
                self._log(f'Put it in this folder:\n{os.getcwd()}\nand try again')
                return
            else:
                home_config_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(service_account_local), str(service_account_correct_location))
                self._log(f"service_account.json has been successfully moved to {home_config_dir}")

            self.config["General"]["installed"] = str(1)
            self._update_config()
            self._log("Restart the App to use it.")
            self.install_button.config(state="disabled")

        except Exception as e:
            error_msg = f"❌ ERROR: {e}\n{traceback.format_exc()}"
            self._log(error_msg)


if __name__ == "__main__":
    App().mainloop()
