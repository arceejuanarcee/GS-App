import tkinter as tk
from tkinter import ttk, messagebox
import requests
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Patch  # <-- for legend boxes


# ---------------------------
# NOAA URLs
# ---------------------------
FORECAST_URL = "https://services.swpc.noaa.gov/products/noaa-planetary-k-index-forecast.json"
NOAA_YEARLY_URL = "https://services.swpc.noaa.gov/text/daily-geomagnetic-indices-{}.txt"


# ---------------------------
# Kp → color mapping (G-scale style)
# ---------------------------
def kp_to_color(kp):
    """
    Approximate SWPC-style G-scale colors.

    Quiet        : Kp < 5       → green
    G1 (Minor)   : 5  ≤ Kp < 6  → yellow
    G2 (Moderate): 6  ≤ Kp < 7  → orange
    G3 (Strong)  : 7  ≤ Kp < 8  → dark orange
    G4 (Severe)  : 8  ≤ Kp < 9  → red
    G5 (Extreme) : Kp ≥ 9       → dark red
    """
    if kp < 5:
        return "green"
    elif kp < 6:
        return "yellow"
    elif kp < 7:
        return "orange"
    elif kp < 8:
        return "darkorange"
    elif kp < 9:
        return "red"
    else:
        return "darkred"


# ---------------------------
# Parse yearly NOAA historical Kp file
# ---------------------------
def fetch_historical_kp(year):
    url = NOAA_YEARLY_URL.format(year)
    print("Fetching:", url)

    r = requests.get(url)
    r.raise_for_status()

    lines = r.text.splitlines()
    kp_data = []

    for line in lines:
        if not line.strip() or line.startswith("#"):
            continue

        parts = line.split()
        if len(parts) < 10:
            continue

        try:
            date = datetime.strptime(parts[0], "%Y-%m-%d")
        except Exception:
            continue

        # Eight 3-hour Kp values
        kp_values = parts[1:9]

        for i, kp in enumerate(kp_values):
            try:
                kp_float = float(kp)
            except Exception:
                continue

            # Time slot = date + (i * 3 hours)
            time_stamp = date.replace(hour=i * 3)
            kp_data.append((time_stamp, kp_float))

    return kp_data


# ---------------------------
# Fetch NOAA 3-day forecast
# ---------------------------
def fetch_forecast():
    r = requests.get(FORECAST_URL)
    r.raise_for_status()
    table = r.json()

    header = table[0]
    rows = table[1:]

    time_i = header.index("time_tag")
    kp_i = header.index("kp")

    data = []
    for row in rows:
        try:
            t = datetime.strptime(row[time_i], "%Y-%m-%d %H:%M:%S")
            kp = float(row[kp_i])
            data.append((t, kp))
        except Exception:
            continue

    return data


# ---------------------------
# GUI App
# ---------------------------
class KPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Geomagnetic Storm Monitor (Kp Index)")
        self.root.geometry("1000x600")

        self.create_widgets()
        self.create_plot()

    def create_widgets(self):
        frame = ttk.Frame(self.root)
        frame.pack(pady=10)

        ttk.Label(frame, text="Mode:").grid(row=0, column=0, padx=5)

        self.mode = ttk.Combobox(
            frame,
            values=["NOAA 3-day Forecast", "Historical (Select Year)"],
            width=30,
        )
        self.mode.grid(row=0, column=1)
        self.mode.bind("<<ComboboxSelected>>", self.toggle_year_selector)
        self.mode.current(0)

        self.year_box = ttk.Combobox(frame, width=10, state="disabled")
        self.year_box.grid(row=0, column=2, padx=10)

        # Fill 20 years of historical options
        current_year = datetime.utcnow().year
        years = [str(y) for y in range(current_year, current_year - 21, -1)]
        self.year_box["values"] = years

        ttk.Button(frame, text="Fetch & Plot", command=self.fetch_and_plot).grid(
            row=0, column=3, padx=10
        )

    def toggle_year_selector(self, event):
        if self.mode.get() == "Historical (Select Year)":
            self.year_box["state"] = "readonly"
            self.year_box.current(0)
        else:
            self.year_box["state"] = "disabled"

    def create_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(10, 4), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def fetch_and_plot(self):
        try:
            if self.mode.get() == "NOAA 3-day Forecast":
                data = fetch_forecast()
                title = "NOAA 3-Day Kp Forecast"
            else:
                year = self.year_box.get()
                if not year:
                    messagebox.showerror("Error", "Select a year first.")
                    return
                data = fetch_historical_kp(year)
                title = f"Historical Kp Data ({year})"

            if not data:
                messagebox.showinfo("No Data", "No Kp data available.")
                return

            # Sort by time
            data.sort(key=lambda x: x[0])
            times = [t for t, _ in data]
            kps = [kp for _, kp in data]

            # ---- COLOR-CODED PLOT ----
            self.ax.clear()

            bar_colors = [kp_to_color(v) for v in kps]
            self.ax.bar(times, kps, width=0.1, color=bar_colors, edgecolor="black", linewidth=0.3)

            self.ax.set_title(title)
            self.ax.set_ylabel("Kp Index")
            self.ax.set_ylim(0, 9)
            self.ax.grid(True, axis="y", linestyle="--", alpha=0.4)

            # Storm threshold line (Kp = 5)
            self.ax.axhline(5, color="red", linestyle="--", linewidth=1)

            # Legend (like NOAA color bar)
            legend_items = [
                ("Kp < 5 (Quiet)", "green"),
                ("Kp = 5 (G1)", "yellow"),
                ("Kp = 6 (G2)", "orange"),
                ("Kp = 7 (G3)", "darkorange"),
                ("Kp ≥ 8 (G4–G5)", "red"),
            ]
            handles = [
                Patch(facecolor=c, edgecolor="black", label=lbl)
                for (lbl, c) in legend_items
            ]
            self.ax.legend(handles=handles, loc="upper right", fontsize=8, framealpha=0.9)

            self.fig.autofmt_xdate()
            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch data:\n{e}")


# ---------------------------
# Run application
# ---------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = KPApp(root)
    root.mainloop()
