import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import gspread
from google.oauth2.service_account import Credentials

# ===================== CONFIGURATION =====================

SERVICE_ACCOUNT_FILE = "C:\\Users\\user-307E6E3400\\Desktop\\Python Scripts\\GS App\\keys\\endless-theorem-421101-fe0721f63c55.json"
SPREADSHEET_ID = "1iR49Cx05EWtbG__o_-gl0SXvHgg5qNBhJ04q7HZN5dQ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# =========================================================

class SheetColumnFetcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Sheets Column B Fetcher")

        # Initialize client & spreadsheet (set in self.init_gspread)
        self.client = None
        self.spreadsheet = None

        # GUI widgets
        self.sheet_label = ttk.Label(root, text="Select Sheet:")
        self.sheet_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(root, textvariable=self.sheet_var, state="readonly", width=40)
        self.sheet_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="we")

        self.reload_button = ttk.Button(root, text="Reload Sheets", command=self.load_sheets)
        self.reload_button.grid(row=0, column=2, padx=5, pady=5)

        self.fetch_button = ttk.Button(root, text="Fetch Column B", command=self.fetch_column_b)
        self.fetch_button.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="we")

        self.text_box = tk.Text(root, wrap="none", width=80, height=20)
        self.text_box.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        self.scrollbar_y = ttk.Scrollbar(root, orient="vertical", command=self.text_box.yview)
        self.scrollbar_y.grid(row=2, column=3, sticky="ns")
        self.text_box.configure(yscrollcommand=self.scrollbar_y.set)

        self.scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=self.text_box.xview)
        self.scrollbar_x.grid(row=3, column=0, columnspan=3, sticky="we")
        self.text_box.configure(xscrollcommand=self.scrollbar_x.set)

        self.save_button = ttk.Button(root, text="Save Output as .txt", command=self.save_to_txt)
        self.save_button.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="we")

        # Make the GUI resize nicely
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Initialize gspread and load sheets
        try:
            self.init_gspread()
            self.load_sheets()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize Google Sheets client:\n{e}")

    def init_gspread(self):
        """Initialize gspread client using service account credentials."""
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)

    def load_sheets(self):
        """Load worksheet names into the dropdown."""
        if self.spreadsheet is None:
            messagebox.showerror("Error", "Spreadsheet is not initialized.")
            return

        try:
            worksheets = self.spreadsheet.worksheets()
            sheet_names = [ws.title for ws in worksheets]
            self.sheet_dropdown["values"] = sheet_names

            if sheet_names:
                # Select first sheet by default
                self.sheet_dropdown.current(0)
            else:
                messagebox.showwarning("No Sheets", "This spreadsheet has no worksheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheets:\n{e}")

    def fetch_column_b(self):
        """Fetch column B from the selected worksheet and display it in the text box."""
        sheet_name = self.sheet_var.get()
        if not sheet_name:
            messagebox.showwarning("Select Sheet", "Please select a sheet first.")
            return

        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
            # Column B is index 2
            column_b_values = worksheet.col_values(2)

            # Clear the text box
            self.text_box.delete("1.0", tk.END)

            # Join each cell in a new line
            output_text = "\n".join(column_b_values)
            self.text_box.insert(tk.END, output_text)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch Column B:\n{e}")

    def save_to_txt(self):
        """Save the content of the text box to a .txt file."""
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
                messagebox.showinfo("Saved", f"File saved successfully:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{e}")


def main():
    root = tk.Tk()
    app = SheetColumnFetcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
