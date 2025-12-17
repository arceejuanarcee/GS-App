import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import numpy as np
import requests
import plotly.graph_objects as go
from plotly.offline import plot as plotly_plot
from skyfield.api import load, EarthSatellite, wgs84

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QSpinBox,
    QMessageBox, QCheckBox, QGroupBox, QFormLayout
)

from PySide6.QtWebEngineWidgets import QWebEngineView


# ==========================================================
# MANUAL SPACE-TRACK CREDENTIALS (KEEP PRIVATE)
# ==========================================================
SPACETRACK_USER = "arceetaraguajuan@gmail.com"
SPACETRACK_PASS = "ArceeJuan123456789"
# ==========================================================


SPACE_TRACK_LOGIN_URL = "https://www.space-track.org/ajaxauth/login"
SPACE_TRACK_QUERY_BASE = "https://www.space-track.org/basicspacedata/query"

SAT_IDS = [
    "43619", "27424", "31698", "37849", "54234", "43013",
    "14780", "25682", "39084", "49260", "43672", "23710",
    "43678", "25994"
]

GROUND_STATION = {
    "name": "Davao GRS",
    "lat": 7.1907,
    "lon": 125.4553
}


@dataclass
class TLE:
    norad: str
    name: str
    line1: str
    line2: str


# ----------------------------------------------------------
# SPACE-TRACK FUNCTIONS
# ----------------------------------------------------------
def spacetrack_login_session(timeout: int = 25) -> requests.Session:
    if not SPACETRACK_USER or not SPACETRACK_PASS:
        raise RuntimeError("Space-Track credentials not set in script.")

    sess = requests.Session()
    r = sess.post(
        SPACE_TRACK_LOGIN_URL,
        data={"identity": SPACETRACK_USER, "password": SPACETRACK_PASS},
        timeout=timeout
    )
    r.raise_for_status()
    return sess


def parse_3le(text: str, fallback_name: str) -> Optional[tuple]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if len(lines) >= 3 and lines[1].startswith("1 ") and lines[2].startswith("2 "):
        return lines[0], lines[1], lines[2]
    if len(lines) >= 2 and lines[0].startswith("1 ") and lines[1].startswith("2 "):
        return fallback_name, lines[0], lines[1]
    return None


def fetch_tle_latest(sess: requests.Session, norad: str) -> TLE:
    url = (
        f"{SPACE_TRACK_QUERY_BASE}/class/tle_latest/"
        f"NORAD_CAT_ID/{norad}/ORDINAL/1/EPOCH/%3Enow-30/format/tle"
    )
    r = sess.get(url, timeout=25)
    r.raise_for_status()

    parsed = parse_3le(r.text, fallback_name=f"NORAD {norad}")
    if not parsed:
        raise RuntimeError(f"No TLE returned for NORAD {norad}")

    name, l1, l2 = parsed
    return TLE(norad=norad, name=name, line1=l1, line2=l2)


def fetch_all_tles(norad_ids: List[str]) -> Dict[str, TLE]:
    sess = spacetrack_login_session()
    return {n: fetch_tle_latest(sess, n) for n in norad_ids}


# ----------------------------------------------------------
# MAP / ORBIT RENDERING (FIXED)
# ----------------------------------------------------------
def build_map_html(selected_tles: List[TLE], minutes_window: int = 45) -> str:
    ts = load.timescale()
    t_now = ts.now()

    eph = load("de421.bsp")
    sun = eph["sun"]
    moon = eph["moon"]

    # ✅ FIXED: correct Earth-centered subpoints
    sun_sp = wgs84.subpoint(sun.at(t_now))
    moon_sp = wgs84.subpoint(moon.at(t_now))

    fig = go.Figure()

    # Ground station
    fig.add_trace(go.Scattergeo(
        lon=[GROUND_STATION["lon"]],
        lat=[GROUND_STATION["lat"]],
        mode="markers+text",
        marker=dict(size=10, symbol="square"),
        text=[GROUND_STATION["name"]],
        textposition="top right",
        name="Ground Station"
    ))

    # Sun
    fig.add_trace(go.Scattergeo(
        lon=[sun_sp.longitude.degrees],
        lat=[sun_sp.latitude.degrees],
        mode="markers+text",
        marker=dict(size=10, color="yellow"),
        text=["Sun"],
        name="Sun"
    ))

    # Moon
    fig.add_trace(go.Scattergeo(
        lon=[moon_sp.longitude.degrees],
        lat=[moon_sp.latitude.degrees],
        mode="markers+text",
        marker=dict(size=10, color="white"),
        text=["Moon"],
        name="Moon"
    ))

    # Track window
    minutes = np.arange(-minutes_window, minutes_window + 1, 1)
    times = ts.utc([
        (datetime.utcnow() + timedelta(minutes=int(m))).year for m in minutes
    ], [
        (datetime.utcnow() + timedelta(minutes=int(m))).month for m in minutes
    ], [
        (datetime.utcnow() + timedelta(minutes=int(m))).day for m in minutes
    ], [
        (datetime.utcnow() + timedelta(minutes=int(m))).hour for m in minutes
    ], [
        (datetime.utcnow() + timedelta(minutes=int(m))).minute for m in minutes
    ], [
        (datetime.utcnow() + timedelta(minutes=int(m))).second for m in minutes
    ])

    for tle in selected_tles:
        sat = EarthSatellite(tle.line1, tle.line2, tle.name, ts)

        sp = sat.at(times).subpoint()
        fig.add_trace(go.Scattergeo(
            lon=sp.longitude.degrees,
            lat=sp.latitude.degrees,
            mode="lines",
            line=dict(width=2),
            name=f"{tle.name} track"
        ))

        sp_now = sat.at(t_now).subpoint()
        fig.add_trace(go.Scattergeo(
            lon=[sp_now.longitude.degrees],
            lat=[sp_now.latitude.degrees],
            mode="markers+text",
            marker=dict(size=9),
            text=[tle.name],
            textposition="top center",
            name=f"{tle.name} now"
        ))

    fig.update_layout(
        geo=dict(
            projection_type="natural earth",
            showland=True,
            showocean=True,
            showcountries=True,
            showcoastlines=True,
        ),
        margin=dict(l=0, r=0, t=40, b=0),
        title=f"Space-Track Ground Track Viewer | UTC {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}",
        legend=dict(y=0.98),
    )

    return plotly_plot(fig, output_type="div", include_plotlyjs="cdn")


# ----------------------------------------------------------
# GUI
# ----------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Satellite Ground Track Viewer (Space-Track Auto)")

        self.tle_store: Dict[str, TLE] = {}
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.auto_refresh)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)

        # LEFT
        left = QVBoxLayout()
        layout.addLayout(left, 0)

        controls = QGroupBox("Controls")
        form = QFormLayout(controls)

        self.btn_fetch = QPushButton("Fetch TLEs")
        self.btn_fetch.clicked.connect(self.fetch_tles)

        self.track_window = QSpinBox()
        self.track_window.setRange(5, 180)
        self.track_window.setValue(45)

        self.btn_render = QPushButton("Render Selected")
        self.btn_render.clicked.connect(self.render_selected)

        self.auto_cb = QCheckBox("Auto-refresh")
        self.refresh_sec = QSpinBox()
        self.refresh_sec.setRange(5, 600)
        self.refresh_sec.setValue(30)

        form.addRow(self.btn_fetch)
        form.addRow("Track window (min):", self.track_window)
        form.addRow(self.btn_render)
        form.addRow(self.auto_cb)
        form.addRow("Refresh every (sec):", self.refresh_sec)

        left.addWidget(controls)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        left.addWidget(QLabel("Select satellites to plot:"))
        left.addWidget(self.list_widget, 1)

        self.status = QLabel("Ready.")
        left.addWidget(self.status)

        # RIGHT
        self.web = QWebEngineView()
        layout.addWidget(self.web, 1)

        self.resize(1400, 800)

    def fetch_tles(self):
        try:
            self.status.setText("Fetching TLEs from Space-Track...")
            self.tle_store = fetch_all_tles(SAT_IDS)
            self.list_widget.clear()

            for tle in self.tle_store.values():
                item = QListWidgetItem(f"{tle.name} (NORAD {tle.norad})")
                item.setData(Qt.UserRole, tle.norad)
                item.setSelected(True)
                self.list_widget.addItem(item)

            self.status.setText("TLEs loaded.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def render_selected(self):
        selected = self.list_widget.selectedItems()
        if not selected:
            return

        tles = [self.tle_store[it.data(Qt.UserRole)] for it in selected]
        html = build_map_html(tles, self.track_window.value())

        page = f"""
        <html><body style="margin:0">
        {html}
        </body></html>
        """
        self.web.setHtml(page)

        if self.auto_cb.isChecked():
            self.timer.start(self.refresh_sec.value() * 1000)
        else:
            self.timer.stop()

    def auto_refresh(self):
        self.fetch_tles()
        self.render_selected()


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
