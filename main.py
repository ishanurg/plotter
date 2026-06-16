#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import math
import numpy as np
import pandas as pd

import pyqtgraph as pg
from pyqtgraph import QtCore, QtGui, QtWidgets

# -------------------- Force GPU acceleration --------------------
pg.setConfigOption('useOpenGL', True)
pg.setConfigOptions(antialias=False)

DARK_STYLE = """
QMainWindow, QDialog, QTabWidget, QWidget { background-color: #1e1e2e; color: #cdd6f4; }
QTabBar::tab { background: #181825; color: #a6adc8; padding: 10px 20px; font-weight: bold; font-size: 14px; border: none; }
QTabBar::tab:selected { background: #313244; color: #89b4fa; border-bottom: 3px solid #89b4fa; }
QTableView { background-color: #1e1e2e; color: #cdd6f4; gridline-color: #313244; border: none; font-size: 13px; }
QHeaderView::section { background-color: #181825; color: #89b4fa; padding: 5px; border: 1px solid #313244; font-weight: bold; }
QPushButton { background-color: #313244; color: #cdd6f4; border: 1px solid #45475a; padding: 6px 12px; border-radius: 4px; font-weight: bold; }
QPushButton:hover { background-color: #89b4fa; color: #1e1e2e; }
QComboBox, QSpinBox, QDoubleSpinBox { background-color: #181825; color: #cdd6f4; border: 1px solid #313244; padding: 4px; border-radius: 3px; }
QLabel { color: #cdd6f4; }
QScrollArea { border: none; }
"""

# ======================================================================
# 1. DATA MODELS
# ======================================================================

class PandasModel(QtCore.QAbstractTableModel):
    """ High-performance model to display Pandas DataFrame in QTableView """
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if index.isValid() and role == QtCore.Qt.DisplayRole:
            val = self._data.iloc[index.row(), index.column()]
            return str(val) if not pd.isna(val) else ""
        return None

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return str(self._data.columns[section])
            if orientation == QtCore.Qt.Vertical:
                return str(self._data.index[section])
        return None

# ======================================================================
# 2. GRAPH COMPONENTS
# ======================================================================

class TurboPlotItem(pg.PlotItem):
    """ Individual Plot instance containing markers, crosshairs, and stats """
    def __init__(self, title, *args, **kwargs):
        super().__init__(title=title, *args, **kwargs)
        self.showGrid(x=True, y=True, alpha=0.15)
        self.setLabel('left', 'Reading', color='#cdd6f4')
        self.setLabel('bottom', 'Relative Time (s)', color='#cdd6f4')
        self.getAxis('left').setPen('#cdd6f4')
        self.getAxis('bottom').setPen('#cdd6f4')
        
        self.getViewBox().setMouseMode(pg.ViewBox.RectMode)

        # Crosshair
        self.vline = pg.InfiniteLine(angle=90, movable=False, pen=pg.mkPen('#f38ba8', width=1))
        self.hline = pg.InfiniteLine(angle=0, movable=False, pen=pg.mkPen('#f38ba8', width=1))
        self.addItem(self.vline, ignoreBounds=True)
        self.addItem(self.hline, ignoreBounds=True)

        # Tooltip
        self.tooltip = pg.TextItem('', anchor=(0, 1), color='#1e1e2e', fill='#cdd6f4')
        self.addItem(self.tooltip, ignoreBounds=True)

        self.curves = {}       # id -> PlotDataItem
        self.data_cache = {}   # id -> (t, y)

        # Markers
        self.marker_mode = False
        self.marker1 = None
        self.marker2 = None
        self.marker_line = pg.PlotDataItem(pen=pg.mkPen('#f9e2af', width=2, style=QtCore.Qt.DashLine))
        self.marker_line.setZValue(100)
        self.addItem(self.marker_line)
        self.marker_label = pg.TextItem('', anchor=(0, 0), color='#1e1e2e', fill='#f9e2af')
        self.marker_label.setZValue(101)
        self.addItem(self.marker_label)
        self.marker_scatter = pg.ScatterPlotItem(brush='#f9e2af', size=10, pen='#f9e2af')
        self.marker_scatter.setZValue(99)
        self.addItem(self.marker_scatter)

        self.scene().sigMouseClicked.connect(self._on_click)
        self.proxy = pg.SignalProxy(self.scene().sigMouseMoved, rateLimit=20, slot=self._mouse_moved)

    def set_marker_mode(self, enabled):
        self.marker_mode = enabled
        self.getViewBox().setMouseMode(pg.ViewBox.PanMode if enabled else pg.ViewBox.RectMode)
        if not enabled:
            self.clear_markers()

    def clear_markers(self):
        self.marker1 = self.marker2 = None
        self.marker_line.setData([], [])
        self.marker_label.setText('')
        self.marker_scatter.setData([], [])

    def _on_click(self, event):
        if not self.marker_mode or event.button() != QtCore.Qt.LeftButton: return
        pos = self.getViewBox().mapSceneToView(event.scenePos())
        x, y = pos.x(), pos.y()

        if self.marker1 is not None and self.marker2 is not None:
            self.marker1 = self.marker2 = None

        if self.marker1 is None: self.marker1 = (x, y)
        elif self.marker2 is None: self.marker2 = (x, y)

        self._update_markers()

    def _update_markers(self):
        pts = [p for p in (self.marker1, self.marker2) if p is not None]
        if not pts:
            self.clear_markers()
            return
            
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        self.marker_scatter.setData(xs, ys)

        if len(pts) == 2:
            self.marker_line.setData(xs, ys)
            dx, dy = xs[1] - xs[0], ys[1] - ys[0]
            slope = dy / dx if dx != 0 else float('inf')
            angle_rad = math.atan(slope) if dx != 0 else (math.pi/2 if dy>0 else -math.pi/2)
            
            txt = f"Δx: {dx:.4g} | Δy: {dy:.4g}\nSlope: {slope:.4g} | Ang: {math.degrees(angle_rad):.2f}°"
            self.marker_label.setText(txt)
            self.marker_label.setPos((xs[0]+xs[1])/2, (ys[0]+ys[1])/2)

    def _mouse_moved(self, evt):
        pos = evt[0]
        if not self.sceneBoundingRect().contains(pos): return
        mouse = self.getViewBox().mapSceneToView(pos)
        x, y = mouse.x(), mouse.y()
        self.vline.setPos(x)
        self.hline.setPos(y)

        best, min_d = None, np.inf
        for ch_id, curve in self.curves.items():
            if not curve.isVisible(): continue
            data = curve.getData()
            if data is None or len(data[0]) == 0: continue
            t_arr, v_arr = data[0], data[1]
            idx = min(np.searchsorted(t_arr, x), len(t_arr)-1)
            d = abs(t_arr[idx] - x)
            if d < min_d:
                min_d, best = d, (ch_id, v_arr[idx])

        if best:
            ch, val = best
            self.tooltip.setText(f"{ch}: t={x:.4f}, v={val:.4g}")
            vr = self.getViewBox().viewRange()
            self.tooltip.setPos(x, y + 0.05 * (vr[1][1] - vr[1][0]))


# ======================================================================
# 3. MAIN APPLICATION TABS
# ======================================================================

class GraphTab(QtWidgets.QWidget):
    def __init__(self, parent_app):
        super().__init__()
        self.app = parent_app
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)

        # -- Left Sidebar (Controls) --
        sidebar = QtWidgets.QWidget()
        sidebar.setFixedWidth(300)
        sidebar_layout = QtWidgets.QVBoxLayout(sidebar)
        
        self.btn_load = QtWidgets.QPushButton("📂 Load CSV Data")
        self.btn_load.clicked.connect(self.app.load_csv)
        sidebar_layout.addWidget(self.btn_load)

        # Layout selection
        layout_box = QtWidgets.QGroupBox("Graph Layout")
        lay_l = QtWidgets.QVBoxLayout(layout_box)
        self.combo_layout = QtWidgets.QComboBox()
        self.combo_layout.addItems(["1 Graph", "2 Graphs (Vertical)", "2 Graphs (Horizontal)", "4 Graphs (2x2)", "6 Graphs (3x2)"])
        self.combo_layout.currentIndexChanged.connect(self.update_grid)
        lay_l.addWidget(self.combo_layout)
        sidebar_layout.addWidget(layout_box)
        
        # Tools
        tools_box = QtWidgets.QGroupBox("Tools")
        t_l = QtWidgets.QVBoxLayout(tools_box)
        self.btn_marker = QtWidgets.QPushButton("📍 Toggle Marker Mode")
        self.btn_marker.setCheckable(True)
        self.btn_marker.clicked.connect(self.toggle_markers)
        t_l.addWidget(self.btn_marker)
        
        btn_reset = QtWidgets.QPushButton("↺ Reset All Views")
        btn_reset.clicked.connect(self.reset_views)
        t_l.addWidget(btn_reset)
        sidebar_layout.addWidget(tools_box)

        # Channel Controller Area
        self.channel_scroll = QtWidgets.QScrollArea()
        self.channel_scroll.setWidgetResizable(True)
        self.channel_container = QtWidgets.QWidget()
        self.channel_layout = QtWidgets.QVBoxLayout(self.channel_container)
        self.channel_layout.setAlignment(QtCore.Qt.AlignTop)
        self.channel_scroll.setWidget(self.channel_container)
        
        sidebar_layout.addWidget(QtWidgets.QLabel("<b>Loaded Channels (Edit Visuals)</b>"))
        sidebar_layout.addWidget(self.channel_scroll)

        layout.addWidget(sidebar)

        # -- Right Area (Graphs) --
        self.graph_layout = pg.GraphicsLayoutWidget()
        self.graph_layout.setBackground("#181825")
        layout.addWidget(self.graph_layout)

        # Initialize maximum 6 subplots
        self.plots = [TurboPlotItem(title=f"Graph {i+1}") for i in range(6)]
        self.update_grid()

    def update_grid(self):
        self.graph_layout.clear()
        idx = self.combo_layout.currentIndex()
        
        layouts = {
            0: [(0,0)],                                       # 1 Graph
            1: [(0,0), (1,0)],                                # 2 Vert
            2: [(0,0), (0,1)],                                # 2 Horz
            3: [(0,0), (0,1), (1,0), (1,1)],                  # 4 Grid
            4: [(0,0), (0,1), (0,2), (1,0), (1,1), (1,2)]     # 6 Grid
        }
        
        positions = layouts[idx]
        for i, pos in enumerate(positions):
            self.graph_layout.addItem(self.plots[i], row=pos[0], col=pos[1])

    def build_channel_controls(self, channels):
        # Clear existing controls
        for i in reversed(range(self.channel_layout.count())): 
            self.channel_layout.itemAt(i).widget().setParent(None)

        palette = ['#89b4fa','#a6e3a1','#f9e2af','#fab387','#f38ba8','#cba6f7']
        self.channel_widgets = {}

        for i, ch_id in enumerate(channels.keys()):
            frame = QtWidgets.QFrame()
            frame.setStyleSheet("QFrame { background-color: #313244; border-radius: 5px; margin-bottom: 5px; padding: 5px; }")
            flay = QtWidgets.QVBoxLayout(frame)
            flay.setContentsMargins(5,5,5,5)

            # Title & Target Graph
            row1 = QtWidgets.QHBoxLayout()
            row1.addWidget(QtWidgets.QLabel(f"<b>{ch_id}</b>"))
            
            target_combo = QtWidgets.QComboBox()
            target_combo.addItems([f"Graph {j+1}" for j in range(6)])
            target_combo.setCurrentIndex(0) # Default all to Graph 1
            target_combo.currentIndexChanged.connect(lambda _, c=ch_id: self.route_data())
            row1.addWidget(target_combo)
            flay.addLayout(row1)

            # Color & Thickness
            row2 = QtWidgets.QHBoxLayout()
            btn_color = QtWidgets.QPushButton("")
            col = palette[i % len(palette)]
            btn_color.setStyleSheet(f"background-color: {col}; border: none; width: 20px; height: 20px;")
            btn_color.clicked.connect(lambda _, c=ch_id, b=btn_color: self.pick_color(c, b))
            
            spin_thick = QtWidgets.QDoubleSpinBox()
            spin_thick.setRange(0.5, 5.0)
            spin_thick.setSingleStep(0.5)
            spin_thick.setValue(1.5)
            spin_thick.valueChanged.connect(lambda _, c=ch_id: self.update_visuals(c))
            
            row2.addWidget(QtWidgets.QLabel("Color:"))
            row2.addWidget(btn_color)
            row2.addWidget(QtWidgets.QLabel("Thickness:"))
            row2.addWidget(spin_thick)
            flay.addLayout(row2)

            self.channel_layout.addWidget(frame)
            
            # Store state
            self.channel_widgets[ch_id] = {
                'color': col,
                'thickness': spin_thick,
                'target': target_combo,
                'color_btn': btn_color
            }

    def pick_color(self, ch_id, btn):
        color = QtWidgets.QColorDialog.getColor()
        if color.isValid():
            hex_col = color.name()
            btn.setStyleSheet(f"background-color: {hex_col}; border: none;")
            self.channel_widgets[ch_id]['color'] = hex_col
            self.update_visuals(ch_id)

    def update_visuals(self, ch_id):
        conf = self.channel_widgets[ch_id]
        pen = pg.mkPen(color=conf['color'], width=conf['thickness'].value(), cosmetic=True)
        
        # Update pen on all plots where this curve exists
        for plot in self.plots:
            if ch_id in plot.curves:
                plot.curves[ch_id].setPen(pen)

    def route_data(self):
        # Clear all existing curves from all plots
        for plot in self.plots:
            for curve in plot.curves.values():
                plot.removeItem(curve)
            plot.curves.clear()

        # Reassign based on target combos
        for ch_id, (t, y) in self.app.loaded_data.items():
            conf = self.channel_widgets[ch_id]
            target_idx = conf['target'].currentIndex()
            target_plot = self.plots[target_idx]
            
            pen = pg.mkPen(color=conf['color'], width=conf['thickness'].value(), cosmetic=True)
            curve = pg.PlotDataItem(t, y, pen=pen, name=str(ch_id))
            target_plot.addItem(curve)
            target_plot.curves[ch_id] = curve

        self.reset_views()

    def toggle_markers(self, checked):
        for plot in self.plots: plot.setMarkerMode(checked)

    def reset_views(self):
        for plot in self.plots: plot.autoRange()

class DataTab(QtWidgets.QWidget):
    def __init__(self, parent_app):
        super().__init__()
        self.app = parent_app
        layout = QtWidgets.QVBoxLayout(self)

        # Toolbar
        toolbar = QtWidgets.QHBoxLayout()
        btn_load = QtWidgets.QPushButton("📂 Load CSV Data")
        btn_load.clicked.connect(self.app.load_csv)
        toolbar.addWidget(btn_load)
        
        self.lbl_info = QtWidgets.QLabel("No data loaded.")
        toolbar.addWidget(self.lbl_info)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Data Table
        self.table = QtWidgets.QTableView()
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        layout.addWidget(self.table)

    def display_data(self, df):
        self.model = PandasModel(df)
        self.table.setModel(self.model)
        self.lbl_info.setText(f"Displaying {df.shape[0]:,} rows and {df.shape[1]} columns.")

# ======================================================================
# 4. MAIN WINDOW
# ======================================================================

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DataBlade Pro – Unified Engine")
        self.resize(1400, 850)
        self.setStyleSheet(DARK_STYLE)

        self.loaded_data = {} # ch_id -> (t, y)
        self.raw_df = None

        # Tabs
        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)

        self.graph_tab = GraphTab(self)
        self.data_tab = DataTab(self)

        self.tabs.addTab(self.graph_tab, "📈 Graph Dashboard")
        self.tabs.addTab(self.data_tab, "📊 Data Analyzer")

        self.statusBar().showMessage("Ready. Load a CSV to begin.")

    def load_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open CSV", "", "CSV Files (*.csv);;All Files (*)")
        if not path: return

        self.statusBar().showMessage("Reading CSV...")
        QtWidgets.QApplication.processEvents()

        try:
            # 1. Load Data
            df = pd.read_csv(path, engine='python', on_bad_lines='skip')
            df.columns = [str(c).strip() for c in df.columns]
            self.raw_df = df

            t_col = next((c for c in df.columns if 'time' in c.lower()), None)
            v_col = next((c for c in df.columns if 'reading' in c.lower() or 'value' in c.lower()), None)
            ch_col = next((c for c in df.columns if 'channel' in c.lower()), None)

            if not t_col or not v_col:
                raise ValueError("Could not auto-detect 'Time' and 'Reading' columns.")

            df_clean = df.dropna(subset=[t_col, v_col])
            
            # 2. Parse Channels
            self.loaded_data.clear()
            groups = df_clean.groupby(ch_col) if ch_col else df_clean.groupby(lambda _: "Signal 1")
            
            for ch, grp in groups:
                t = pd.to_numeric(grp[t_col], errors='coerce').dropna().values.astype(np.float32)
                y = pd.to_numeric(grp[v_col], errors='coerce').dropna().values.astype(np.float32)
                # Sort by time
                idx = np.argsort(t)
                self.loaded_data[str(ch)] = (t[idx], y[idx])

            # 3. Update UI
            self.data_tab.display_data(self.raw_df)
            self.graph_tab.build_channel_controls(self.loaded_data)
            self.graph_tab.route_data() # Distribute data to plots
            
            total_pts = sum(len(y) for t, y in self.loaded_data.values())
            self.statusBar().showMessage(f"Successfully loaded {total_pts:,} data points.")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load CSV:\n{str(e)}")
            self.statusBar().showMessage("Error loading file.")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    
    # Modern font
    font = QtGui.QFont("Segoe UI", 10)
    app.setFont(font)
    
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())