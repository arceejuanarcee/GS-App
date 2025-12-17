import sys
import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

import numpy as np
import requests
import plotly.graph_objects as go
from plotly.offline import plot as plotly_plot
from skyfield.api import load, EarthSatellite, wgs84

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QSpinBox,
    QMessageBox, QCheckBox, QGroupBox, QFormLayout, QTabWidget
)
from PySide6.QtWebEngineWidgets import QWebEngineView


# ======= Space-Track credentials ===========
SPACETRACK_USER = "arceetaraguajuan@gmail.com"
SPACETRACK_PASS = "ArceeJuan123456789"
# ===========================================

SPACE_TRACK_LOGIN_URL = "https://www.space-track.org/ajaxauth/login"
SPACE_TRACK_QUERY_BASE = "https://www.space-track.org/basicspacedata/query"

SAT_IDS = [
    "43619", "27424", "31698", "37849", "54234", "43013",
    "14780", "25682", "39084", "49260", "43672", "23710",
    "43678", "25994"
]

GROUND_STATION = {"name": "Davao GRS", "lat": 7.1907, "lon": 125.4553}


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
        raise RuntimeError("Space-Track credentials not set in script (SPACETRACK_USER / SPACETRACK_PASS).")

    sess = requests.Session()
    r = sess.post(
        SPACE_TRACK_LOGIN_URL,
        data={"identity": SPACETRACK_USER, "password": SPACETRACK_PASS},
        timeout=timeout
    )
    r.raise_for_status()
    return sess


def parse_3le(text: str, fallback_name: str) -> Optional[Tuple[str, str, str]]:
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
    out: Dict[str, TLE] = {}
    for n in norad_ids:
        out[n] = fetch_tle_latest(sess, n)
    return out


# ----------------------------------------------------------
# MAP (MAPBOX + ESRI SATELLITE IMAGERY) + LIVE HOOKS
# ----------------------------------------------------------
def build_map_html_with_live_hooks_mapbox(
    selected_tles: List[TLE],
    minutes_window: int,
    tail_len: int
) -> Tuple[str, List[dict]]:
    """
    Uses Mapbox with an ESRI World Imagery raster layer (satellite-style Earth).
    Returns:
      - HTML string
      - sat_trace_meta (trace indices)
    """
    ts = load.timescale()
    t_now = ts.now()

    eph = load("de421.bsp")
    earth = eph["earth"]
    sun = eph["sun"]
    moon = eph["moon"]

    # Sun/Moon subpoints (Earth-centered)
    sun_geo = earth.at(t_now).observe(sun).apparent()
    moon_geo = earth.at(t_now).observe(moon).apparent()
    sun_sp = wgs84.subpoint(sun_geo)
    moon_sp = wgs84.subpoint(moon_geo)

    fig = go.Figure()

    # Track times (static)
    base = datetime.utcnow()
    mins = np.arange(-minutes_window, minutes_window + 1, 1, dtype=int)
    dt_list = [base + timedelta(minutes=int(m)) for m in mins]
    times = ts.utc(
        [d.year for d in dt_list],
        [d.month for d in dt_list],
        [d.day for d in dt_list],
        [d.hour for d in dt_list],
        [d.minute for d in dt_list],
        [d.second for d in dt_list],
    )

    # Ground Station, Sun, Moon as scattermapbox
    fig.add_trace(go.Scattermapbox(
        lon=[GROUND_STATION["lon"]],
        lat=[GROUND_STATION["lat"]],
        mode="markers+text",
        marker=dict(size=12),
        text=[GROUND_STATION["name"]],
        textposition="top right",
        name="Ground Station"
    ))

    fig.add_trace(go.Scattermapbox(
        lon=[sun_sp.longitude.degrees],
        lat=[sun_sp.latitude.degrees],
        mode="markers+text",
        marker=dict(size=12),
        text=["Sun"],
        textposition="top center",
        name="Sun"
    ))

    fig.add_trace(go.Scattermapbox(
        lon=[moon_sp.longitude.degrees],
        lat=[moon_sp.latitude.degrees],
        mode="markers+text",
        marker=dict(size=12),
        text=["Moon"],
        textposition="top center",
        name="Moon"
    ))

    sat_trace_meta: List[dict] = []

    # Add satellites: static track + live tail + live marker
    for tle in selected_tles:
        sat = EarthSatellite(tle.line1, tle.line2, tle.name, ts)

        sp_track = sat.at(times).subpoint()
        track_lons = sp_track.longitude.degrees
        track_lats = sp_track.latitude.degrees

        track_i = len(fig.data)
        fig.add_trace(go.Scattermapbox(
            lon=track_lons.tolist(),
            lat=track_lats.tolist(),
            mode="lines",
            line=dict(width=2),
            name=f"{tle.name} track"
        ))

        sp_now = sat.at(t_now).subpoint()
        cur_lon = float(sp_now.longitude.degrees)
        cur_lat = float(sp_now.latitude.degrees)

        tail_i = len(fig.data)
        fig.add_trace(go.Scattermapbox(
            lon=[cur_lon],
            lat=[cur_lat],
            mode="lines",
            line=dict(width=3),
            name=f"{tle.name} tail"
        ))

        marker_i = len(fig.data)
        fig.add_trace(go.Scattermapbox(
            lon=[cur_lon],
            lat=[cur_lat],
            mode="markers+text",
            marker=dict(size=10),
            text=[tle.name],
            textposition="top center",
            name=f"{tle.name} now"
        ))

        sat_trace_meta.append({
            "norad": tle.norad,
            "name": tle.name,
            "track_i": track_i,
            "tail_i": tail_i,
            "marker_i": marker_i
        })

    # --- Satellite imagery basemap using an ESRI raster layer ---
    # Mapbox token not required because we use style="white-bg" + custom raster tiles.
    # (Works without you needing to provide a token.)
    fig.update_layout(
        mapbox=dict(
            style="white-bg",
            center=dict(lat=0, lon=0),
            zoom=0.8,
            layers=[
                dict(
                    sourcetype="raster",
                    source=[
                        "https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
                    ],
                    below="traces"
                )
            ],
        ),
        margin=dict(l=0, r=0, t=40, b=0),
        title=f"Live Ground Track Viewer | UTC {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}",
        legend=dict(y=0.98),
    )

    div = plotly_plot(fig, output_type="div", include_plotlyjs="cdn")

    meta_json = json.dumps(sat_trace_meta)
    tail_len_js = int(tail_len)

    page = f"""
    <html>
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Live Satellite Map</title>
        <style>
          body {{ margin:0; font-family: Arial, sans-serif; background: #000; }}
          #wrap {{ width:100vw; height:100vh; }}
        </style>
      </head>
      <body>
        <div id="wrap">{div}</div>

        <script>
          function getPlotDiv() {{
            return document.querySelector('.plotly-graph-div');
          }}

          const SAT_META = {meta_json};
          const TAIL_LEN = {tail_len_js};
          const tailBuffers = {{}};

          function initTails() {{
            SAT_META.forEach(s => {{
              tailBuffers[s.norad] = {{ lons: [], lats: [] }};
            }});
          }}

          function pushTail(norad, lon, lat) {{
            const b = tailBuffers[norad];
            if (!b) return;
            b.lons.push(lon);
            b.lats.push(lat);
            if (b.lons.length > TAIL_LEN) {{
              b.lons = b.lons.slice(-TAIL_LEN);
              b.lats = b.lats.slice(-TAIL_LEN);
            }}
          }}

          window.updateSatellite = function(norad, lon, lat) {{
            const plotDiv = getPlotDiv();
            if (!plotDiv) return;

            if (!tailBuffers[norad]) {{
              tailBuffers[norad] = {{ lons: [], lats: [] }};
            }}
            pushTail(norad, lon, lat);

            const s = SAT_META.find(x => x.norad === norad);
            if (!s) return;

            // Update marker (single point)
            Plotly.restyle(plotDiv, {{ lon: [[lon]], lat: [[lat]] }}, [s.marker_i]);

            // Update tail line
            const b = tailBuffers[norad];
            Plotly.restyle(plotDiv, {{ lon: [b.lons], lat: [b.lats] }}, [s.tail_i]);
          }}

          window.resetTails = function() {{
            initTails();
          }}

          initTails();
        </script>
      </body>
    </html>
    """

    return page, sat_trace_meta


# ----------------------------------------------------------
# GUI
# ----------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Satellite Ground Track Viewer (Space-Track Live)")

        self.tle_store: Dict[str, TLE] = {}
        self.selected_norads: List[str] = []

        self.ts = load.timescale()
        self.sat_objects: Dict[str, EarthSatellite] = {}

        self.live_timer = QTimer(self)
        self.live_timer.timeout.connect(self.live_tick)

        self.web_loaded = False

        # --- Tabs root ---
        tabs = QTabWidget()
        self.setCentralWidget(tabs)

        # TAB 1: VIEW (map only)
        view_tab = QWidget()
        view_layout = QVBoxLayout(view_tab)
        self.web = QWebEngineView()
        self.web.loadFinished.connect(self.on_web_load_finished)
        view_layout.addWidget(self.web)
        tabs.addTab(view_tab, "View")

        # TAB 2: CONTROLS (no map)
        controls_tab = QWidget()
        controls_layout = QVBoxLayout(controls_tab)

        controls_group = QGroupBox("Controls / Settings")
        form = QFormLayout(controls_group)

        self.btn_fetch = QPushButton("Fetch TLEs (SAT_IDS)")
        self.btn_fetch.clicked.connect(self.fetch_tles)

        self.track_window = QSpinBox()
        self.track_window.setRange(5, 180)
        self.track_window.setValue(45)

        self.tail_len = QSpinBox()
        self.tail_len.setRange(5, 600)
        self.tail_len.setValue(90)

        self.btn_render = QPushButton("Render Selected (Start Live)")
        self.btn_render.clicked.connect(self.render_selected)

        self.live_cb = QCheckBox("Live update (moving)")
        self.live_cb.setChecked(True)

        self.live_interval = QSpinBox()
        self.live_interval.setRange(200, 5000)
        self.live_interval.setValue(1000)

        form.addRow(self.btn_fetch)
        form.addRow("Track window (min):", self.track_window)
        form.addRow("Tail length (points):", self.tail_len)
        form.addRow(self.btn_render)
        form.addRow(self.live_cb)
        form.addRow("Live interval (ms):", self.live_interval)

        controls_layout.addWidget(controls_group)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        controls_layout.addWidget(QLabel("Select satellites to plot:"))
        controls_layout.addWidget(self.list_widget, 1)

        self.status = QLabel("Ready.")
        self.status.setWordWrap(True)
        controls_layout.addWidget(self.status)

        tabs.addTab(controls_tab, "Controls")

        self.resize(1400, 800)

    def set_status(self, msg: str):
        self.status.setText(msg)

    def fetch_tles(self):
        try:
            self.set_status("Fetching TLEs from Space-Track...")
            self.tle_store = fetch_all_tles(SAT_IDS)

            self.list_widget.clear()
            for norad, tle in sorted(self.tle_store.items(), key=lambda x: x[1].name.upper()):
                item = QListWidgetItem(f"{tle.name} (NORAD {tle.norad})")
                item.setData(Qt.UserRole, tle.norad)
                self.list_widget.addItem(item)

            for i in range(self.list_widget.count()):
                self.list_widget.item(i).setSelected(True)

            self.set_status("TLEs loaded. Select satellites then click Render Selected.")
        except Exception as e:
            QMessageBox.critical(self, "Fetch failed", str(e))
            self.set_status("Fetch failed.")

    def render_selected(self):
        try:
            selected_items = self.list_widget.selectedItems()
            if not selected_items:
                QMessageBox.information(self, "No selection", "Select one or more satellites first.")
                return
            if not self.tle_store:
                QMessageBox.information(self, "No TLEs", "Click Fetch TLEs first.")
                return

            self.selected_norads = []
            selected_tles: List[TLE] = []
            self.sat_objects.clear()

            for it in selected_items:
                norad = it.data(Qt.UserRole)
                if norad in self.tle_store:
                    tle = self.tle_store[norad]
                    selected_tles.append(tle)
                    self.selected_norads.append(norad)
                    self.sat_objects[norad] = EarthSatellite(tle.line1, tle.line2, tle.name, self.ts)

            self.web_loaded = False
            self.set_status(f"Rendering {len(selected_tles)} satellite(s)...")

            page, _meta = build_map_html_with_live_hooks_mapbox(
                selected_tles=selected_tles,
                minutes_window=int(self.track_window.value()),
                tail_len=int(self.tail_len.value())
            )

            self.web.setHtml(page)
        except Exception as e:
            QMessageBox.critical(self, "Render failed", str(e))
            self.set_status("Render failed.")

    def on_web_load_finished(self, ok: bool):
        self.web_loaded = bool(ok)
        if not ok:
            self.set_status("Map failed to load.")
            return

        self.web.page().runJavaScript("if (window.resetTails) window.resetTails();")

        if self.live_cb.isChecked():
            self.live_timer.start(int(self.live_interval.value()))
            self.set_status("Rendered. Live update started.")
        else:
            self.live_timer.stop()
            self.set_status("Rendered. Live update OFF.")

    def live_tick(self):
        if not self.web_loaded or not self.selected_norads:
            return

        try:
            t_now = self.ts.now()
            for norad in self.selected_norads:
                sat = self.sat_objects.get(norad)
                if sat is None:
                    continue

                sp = sat.at(t_now).subpoint()
                lon = float(sp.longitude.degrees)
                lat = float(sp.latitude.degrees)

                js = f"if (window.updateSatellite) window.updateSatellite({json.dumps(norad)}, {lon}, {lat});"
                self.web.page().runJavaScript(js)

        except Exception as e:
            self.live_timer.stop()
            QMessageBox.critical(self, "Live update failed", str(e))
            self.set_status("Live update failed (stopped).")


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
