import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime
import os

# ========== CONFIG ==========
SERVICE_ACCOUNT_FILE = "C:\\Users\\user-307E6E3400\\Desktop\\Python Scripts\\GS App\\keys\\endless-theorem-421101-fe0721f63c55.json"
SPREADSHEET_ID = "1iR49Cx05EWtbG__o_-gl0SXvHgg5qNBhJ04q7HZN5dQ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
# ============================

# ---------- SC CONVERTER LOGIC (from your Streamlit script) ----------

def parse_command_line(line):
    """
    Parse a single raw command line into (cmd_id, descriptor, params).
    Returns None if the line does not start with a valid 0x?? command.
    """
    m = re.match(r'^\s*0[xX]([0-9A-Fa-f]{2})\s+([0-9A-Fa-f]{2,4})', line)
    if not m:
        return None
    cmd_id = m.group(1).upper()
    code_str = m.group(2).upper()
    rest = line[m.end():]

    # Find descriptor token (skip SC_WAIT_A & SC_DATE)
    tokens = re.findall(r'#([^ \t#]+)', rest)
    desc = None
    for t in tokens:
        if t.startswith('SC_WAIT_A') or t.startswith('SC_DATE'):
            continue
        desc = t
        break
    if desc is None:
        return None

    # Build prefix
    if len(code_str) == 4:
        prefix = f"({code_str[:2]})({code_str[2:]})"
    else:
        prefix = f"({code_str})"
    descriptor = f"{prefix} {desc}"

    # Collect all hex bytes before descriptor
    pre_desc = rest.split(f"#{desc}", 1)[0]
    params = re.findall(r'([0-9A-Fa-f]{2})', pre_desc)
    param_str = ",".join(p.upper() for p in params)

    return cmd_id, descriptor, param_str


def convert_log(file_contents):
    sc_date_re = re.compile(r'#SC_DATE=(\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})')
    wait_re = re.compile(r'#SC_WAIT_A=(\d+)')
    upload_dt = datetime.now()
    last_wait = None
    warnings = []
    output_lines = []
    console_output = ""

    for lineno, raw in enumerate(file_contents.splitlines(), start=1):
        w_match = wait_re.search(raw)
        if w_match:
            last_wait = w_match.group(1)

        sd = sc_date_re.search(raw)
        if sd and last_wait is not None:
            try:
                sc_dt_pst = datetime.strptime(sd.group(1), '%Y/%m/%d %H:%M:%S')
                shifted_hour = (sc_dt_pst.hour + 16) % 24
                shifted = sc_dt_pst.replace(hour=shifted_hour)
                yy = shifted.year % 100
                parts = [yy, shifted.month, shifted.day, shifted.hour, shifted.minute, shifted.second]
                hex_params = ",".join(f"{p:02X}" for p in parts)
                output_lines.append(f"{last_wait}\t3E\t(03)(01) SC_TIME_SET\t{hex_params}")
                if sc_dt_pst < upload_dt:
                    dt_str = sc_dt_pst.strftime("%d/%m/%Y %I:%M:%S %p").lower()
                    dt_str = re.sub(r'\b0([1-9]):', r'\1:', dt_str)
                    console_output += f"[Warning] Line {lineno}\n"
                    console_output += f"    Commands executed on {dt_str} +08:00 already elapsed. Please check.\n"
                    warnings.append(lineno)
            except Exception:
                console_output += f"[Warning] Line {lineno}\n    Invalid SC_DATE format.\n"

        parsed = parse_command_line(raw)
        if parsed and last_wait is not None:
            cmd_id, descriptor, params = parsed
            if params:
                output_lines.append(f"{last_wait}\t{cmd_id}\t{descriptor}\t{params}")
            else:
                output_lines.append(f"{last_wait}\t{cmd_id}\t{descriptor}\t")

    console_output += f"\nSC file converted with {len(warnings)} warning(s).\n"
    return "\n".join(output_lines), console_output

# ---------- MAIN GUI APP ----------

class SheetColumnFetcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Sheets Column B Fetcher + SC Converter")

        self.client = None
        self.spreadsheet = None

        self.sheet_var = tk.StringVar()

        # --- Top row: sheet selection ---
        ttk.Label(root, text="Select Sheet:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sheet_dropdown = ttk.Combobox(root, textvariable=self.sheet_var, state="readonly", width=40)
        self.sheet_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(root, text="Reload Sheets", command=self.load_sheets).grid(row=0, column=2, padx=5, pady=5)

        # --- Fetch button ---
        ttk.Button(root, text="Fetch Column B", command=self.fetch_column_b).grid(
            row=1, column=0, columnspan=3, padx=5, pady=5, sticky="we"
        )

        # --- Text box for fetched content ---
        self.text_box = tk.Text(root, wrap="none", width=80, height=20)
        self.text_box.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        # Scrollbars
        scroll_y = ttk.Scrollbar(root, orient="vertical", command=self.text_box.yview)
        scroll_y.grid(row=2, column=3, sticky="ns")
        self.text_box.configure(yscrollcommand=scroll_y.set)

        scroll_x = ttk.Scrollbar(root, orient="horizontal", command=self.text_box.xview)
        scroll_x.grid(row=3, column=0, columnspan=3, sticky="we")
        self.text_box.configure(xscrollcommand=scroll_x.set)

        # --- Buttons: save + convert ---
        ttk.Button(root, text="Save Output as .txt", command=self.save_to_txt).grid(
            row=4, column=0, columnspan=3, padx=5, pady=5, sticky="we"
        )
        ttk.Button(root, text="Convert SC Commands", command=self.convert_sc_commands).grid(
            row=5, column=0, columnspan=3, padx=5, pady=5, sticky="we"
        )

        # Layout config
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Initialize Sheets
        try:
            self.init_gspread()
            self.load_sheets()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize Google Sheets client:\n{e}")

    # -------- Google Sheets handling --------

    def init_gspread(self):
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)

    def load_sheets(self):
        if not self.spreadsheet:
            return
        try:
            sheets = self.spreadsheet.worksheets()
            names = [s.title for s in sheets]
            self.sheet_dropdown["values"] = names
            if names:
                self.sheet_dropdown.current(0)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheets:\n{e}")

    def fetch_column_b(self):
        sheet_name = self.sheet_var.get()
        if not sheet_name:
            messagebox.showwarning("Select Sheet", "Please select a sheet first.")
            return

        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
            col_b = worksheet.col_values(2)  # Column B

            self.text_box.delete("1.0", tk.END)
            output_text = "\n".join(col_b)
            self.text_box.insert(tk.END, output_text)

            # Auto-save to fetched_column_b.txt in script folder
            script_dir = os.path.dirname(os.path.abspath(__file__))
            auto_path = os.path.join(script_dir, "fetched_column_b.txt")
            with open(auto_path, "w", encoding="utf-8") as f:
                f.write(output_text)

            messagebox.showinfo(
                "Success",
                f"Column B fetched.\nAuto-saved as:\n{auto_path}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch Column B:\n{e}")

    def save_to_txt(self):
        content = self.text_box.get("1.0", tk.END).rstrip("\n")
        if not content.strip():
            messagebox.showwarning("No Content", "There is no content to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Save as"
        )
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(content)
                messagebox.showinfo("Saved", f"File saved:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{e}")

    # -------- SC conversion via GUI --------

    def convert_sc_commands(self):
        raw_text = self.text_box.get("1.0", tk.END).rstrip("\n")
        if not raw_text.strip():
            messagebox.showwarning("No Content", "No text to convert. Fetch Column B or paste SC log first.")
            return

        try:
            converted_text, console_output = convert_log(raw_text)
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed:\n{e}")
            return

        # Create a new window to show results
        win = tk.Toplevel(self.root)
        win.title("SC Commands Conversion Result")
        win.geometry("900x600")

        # Frames for output + console
        frame_top = ttk.LabelFrame(win, text="Converted SC File")
        frame_top.pack(fill="both", expand=True, padx=5, pady=5)

        frame_bottom = ttk.LabelFrame(win, text="Console Output / Warnings")
        frame_bottom.pack(fill="both", expand=True, padx=5, pady=5)

        # Converted text box
        txt_converted = tk.Text(frame_top, wrap="none", height=15)
        txt_converted.pack(fill="both", expand=True, padx=5, pady=5)
        txt_converted.insert(tk.END, converted_text)

        # Console text box
        txt_console = tk.Text(frame_bottom, wrap="none", height=10)
        txt_console.pack(fill="both", expand=True, padx=5, pady=5)
        txt_console.insert(tk.END, console_output.strip())

        # Save button inside the new window
        def save_converted():
            path = filedialog.asksaveasfilename(
                parent=win,
                defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
                title="Save converted SC file as"
            )
            if path:
                try:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(converted_text)
                    messagebox.showinfo("Saved", f"Converted file saved:\n{path}")
                except Exception as e_inner:
                    messagebox.showerror("Error", f"Failed to save converted file:\n{e_inner}")

        btn_save_conv = ttk.Button(win, text="Save Converted File as .txt", command=save_converted)
        btn_save_conv.pack(padx=5, pady=5)


def main():
    root = tk.Tk()
    app = SheetColumnFetcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
