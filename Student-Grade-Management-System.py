"""
VIT-Yarthi Student Grade Management System 

- Robust handling of corrupted students.json (auto-repair)
- All GUI features: Add/Update/Delete, Search, Sort, CSV, PDF
- Charts: Line & Bar (improved visuals, data labels, hover)
- Grading: S/A/B/C/D/E/N (Fail) per requested ranges

Dependencies:
    pip install matplotlib reportlab pillow

"""

import json
import csv
import os
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
import math

# Logo path (optional)
LOGO_PATH = "/mnt/data/e0504538-8166-45e2-8ff5-de5ab1ec01e9.png"

DATA_FILE = "students.json"


# ---------------- Data Manager ----------------
class StudentManager:
    def __init__(self, file_path=DATA_FILE):
        self.file_path = file_path
        self.students = {}  # roll -> {name, marks}
        self.load()

    def load(self):
        # load and auto-repair old simple formats
        if not os.path.exists(self.file_path):
            self.students = {}
            return

        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
        except Exception:
            self.students = {}
            return

        fixed = {}
        for roll, value in raw.items():
            if isinstance(value, dict):
                name = value.get('name', 'UNKNOWN')
                marks = value.get('marks', 0)
                try:
                    marks = float(marks)
                except Exception:
                    marks = 0.0
                fixed[str(roll).upper()] = {"name": str(name).upper(), "marks": marks}
            else:
                try:
                    marks = float(value)
                except Exception:
                    marks = 0.0
                fixed[str(roll).upper()] = {"name": "UNKNOWN", "marks": marks}

        self.students = fixed
        # persist repaired
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self.students, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def save(self):
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self.students, f, indent=2, ensure_ascii=False)
        except Exception:
            # avoid modal dialogs from background threads; show minimal warning
            print("Warning: failed to save data to disk.")

    def add_update(self, roll, name, marks):
        self.students[str(roll).upper()] = {"name": str(name).upper(), "marks": float(marks)}
        self.save()

    def delete(self, roll):
        roll = str(roll).upper()
        if roll in self.students:
            del self.students[roll]
            self.save()

    def export_csv(self, path):
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(["Rank", "Roll", "Name", "Marks", "Grade"])
                ranked = sorted(self.students.items(), key=lambda x: x[1]['marks'], reverse=True)
                for i, (r, s) in enumerate(ranked, 1):
                    w.writerow([i, r, s['name'], s['marks'], calculate_grade(s['marks'])])
        except Exception as e:
            messagebox.showerror("Export error", f"Failed to export CSV:\n{e}")

    def stats(self):
        if not self.students:
            return {"total": 0, "max": 0, "min": 0, "avg": 0}
        marks = [s['marks'] for s in self.students.values()]
        return {"total": len(marks), "max": max(marks), "min": min(marks), "avg": sum(marks) / len(marks)}

    def sorted(self, mode='Roll'):
        if mode == 'Name':
            return dict(sorted(self.students.items(), key=lambda x: x[1]['name'].lower()))
        elif mode == 'Marks':
            return dict(sorted(self.students.items(), key=lambda x: x[1]['marks'], reverse=True))
        else:
            def roll_key(item):
                r = item[0]
                try:
                    return int(r)
                except Exception:
                    return r
            return dict(sorted(self.students.items(), key=roll_key))

    def filter_search(self, text):
        if not text:
            return set(self.students.keys())
        t = text.lower()
        return {r for r, s in self.students.items() if t in r.lower() or t in s['name'].lower()}


# ---------------- Helpers ----------------
def calculate_grade(marks):
    try:
        m = float(marks)
    except Exception:
        m = 0.0
    # per user's specification:
    if m >= 90:
        return 'S'
    if m >= 80:
        return 'A'
    if m >= 70:
        return 'B'
    if m >= 60:
        return 'C'
    if m >= 50:
        return 'D'
    if m >= 40:
        return 'E'
    return 'N'  # fail


# ---------------- App ----------------
class GradeApp:
    def __init__(self, root):
        self.root = root
        self.manager = StudentManager()
        self.dark = True
        self.root.title('VIT-Yarthi Student Grade Management System')
        self.root.geometry('1150x750')
        self.setup_styles()
        self.build_ui()
        # chart interaction state
        self.bar_rects = []
        self.line_points = []  # list of (x, y, roll, name, marks)
        self.hover_ann = None
        self.bar_cid = None
        self.line_cid = None
        # Ensure data shape is correct before first refresh
        self.manager.load()
        self.refresh_all()

    def setup_styles(self):
        self.style = ttk.Style(self.root)
        self.set_theme(self.dark)

    def set_theme(self, dark=True):
        self.dark = dark
        bg = '#0f0f0f' if dark else '#f2f2f2'
        fg = 'white' if dark else 'black'
        entry_bg = '#1b1b1b' if dark else 'white'
        btn_bg = '#2b2b2b' if dark else '#1f6feb'

        self.root.configure(bg=bg)
        self.style.theme_use('clam')
        self.style.configure('TLabel', background=bg, foreground=fg, font=('Segoe UI', 10))
        self.style.configure('Header.TLabel', background=bg, foreground=fg, font=('Segoe UI', 14, 'bold'))
        self.style.configure('TEntry', fieldbackground=entry_bg, foreground=fg)
        self.style.configure('TButton', foreground=fg, background=btn_bg, font=('Segoe UI', 10))
        self.style.map('TButton', background=[('active', '#3a7bd5')])
        self.style.configure('Treeview', background='#222' if dark else 'white', foreground=fg, fieldbackground='#222' if dark else 'white')
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))

    def build_ui(self):
        # Top heading
        top = tk.Frame(self.root, bg=self.root['bg'])
        top.pack(side='top', fill='x', padx=8, pady=6)
        ttk.Label(top, text='VIT-Yarthi Student Grade Management System', style='Header.TLabel').pack()

        # Left controls
        left = tk.Frame(self.root, bg=self.root['bg'])
        left.pack(side='left', fill='y', padx=12, pady=12)

        ttk.Label(left, text='Roll Number:').pack(anchor='w')
        self.roll_var = tk.StringVar()
        e_roll = ttk.Entry(left, textvariable=self.roll_var)
        e_roll.pack(fill='x')

        ttk.Label(left, text='Name:').pack(anchor='w', pady=(8, 0))
        self.name_var = tk.StringVar()
        e_name = ttk.Entry(left, textvariable=self.name_var)
        e_name.pack(fill='x')

        ttk.Label(left, text='Marks (0-100):').pack(anchor='w', pady=(8, 0))
        self.mark_var = tk.StringVar()
        e_mark = ttk.Entry(left, textvariable=self.mark_var)
        e_mark.pack(fill='x')

        # Auto-uppercase while typing for roll and name
        def to_upper_roll(*args):
            v = self.roll_var.get()
            up = v.upper()
            if v != up:
                self.roll_var.set(up)

        def to_upper_name(*args):
            v = self.name_var.get()
            up = v.upper()
            if v != up:
                self.name_var.set(up)

        try:
            self.roll_var.trace_add('write', lambda *a: to_upper_roll())
            self.name_var.trace_add('write', lambda *a: to_upper_name())
        except Exception:
            self.roll_var.trace('w', lambda *a: to_upper_roll())
            self.name_var.trace('w', lambda *a: to_upper_name())
        ttk.Button(left, text='Add / Update', command=self.add_update).pack(fill='x', pady=8)
        ttk.Button(left, text='Delete', command=self.delete).pack(fill='x')
        ttk.Button(left, text='Export CSV', command=self.export_csv).pack(fill='x', pady=8)

        ttk.Separator(left, orient='horizontal').pack(fill='x', pady=8)
        ttk.Label(left, text='Search:').pack(anchor='w')
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(left, textvariable=self.search_var)
        search_entry.pack(fill='x')
        try:
            self.search_var.trace_add('write', lambda *a: self.refresh_table())
        except Exception:
            self.search_var.trace('w', lambda *a: self.refresh_table())

        ttk.Label(left, text='Sort By:').pack(anchor='w', pady=(8, 0))
        self.sort_mode = tk.StringVar(value='Roll')
        ttk.OptionMenu(left, self.sort_mode, 'Roll', 'Roll', 'Name', 'Marks', command=lambda *_: self.refresh_all()).pack(fill='x')

        ttk.Button(left, text='Toggle Light/Dark', command=self.toggle_theme).pack(fill='x', pady=8)
        ttk.Button(left, text='Generate PDF Report', command=self.generate_pdf).pack(fill='x', pady=4)

        # Stats
        self.stats_label = ttk.Label(left, text='')
        self.stats_label.pack(anchor='w', pady=12)

        # Right: table and charts
        right = tk.Frame(self.root, bg=self.root['bg'])
        right.pack(side='right', expand=True, fill='both', padx=12, pady=12)

        # Table
        cols = ('Roll', 'Name', 'Marks', 'Grade')
        self.tree = ttk.Treeview(right, columns=cols, show='headings', height=8, selectmode='browse')
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=220 if c == 'Name' else 110, anchor='center')
        self.tree.pack(fill='x', pady=(0, 8))
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # Charts area stacked vertically - now only two charts: line (top) and bar (bottom)
        charts = tk.Frame(right, bg=self.root['bg'])
        charts.pack(fill='both', expand=True)

        # Create two figures (bigger than before)
        self.fig_line, self.ax_line = plt.subplots(figsize=(10, 3.0), dpi=110)
        self.canvas_line = FigureCanvasTkAgg(self.fig_line, master=charts)
        self.canvas_line.get_tk_widget().pack(fill='both', expand=True, pady=(0, 6))

        self.fig_bar, self.ax_bar = plt.subplots(figsize=(12, 3), dpi=100)
        self.canvas_bar = FigureCanvasTkAgg(self.fig_bar, master=charts)
        self.canvas_bar.get_tk_widget().pack(fill='both', expand=True)

        # Improve layout spacing
        plt.tight_layout()

    # ---------------- CRUD ----------------
    def add_update(self):
        roll = self.roll_var.get().strip().upper()
        name = self.name_var.get().strip().upper()
        mark_s = self.mark_var.get().strip()
        if not roll or not name:
            messagebox.showwarning('Input error', 'Roll and Name are required')
            return
        try:
            marks = float(mark_s)
            if not (0 <= marks <= 100):
                raise ValueError
        except Exception:
            messagebox.showwarning('Input error', 'Enter marks between 0 and 100')
            return
        self.manager.add_update(roll, name, marks)
        self.clear_fields()
        self.refresh_all()

    def delete(self):
        roll = self.roll_var.get().strip().upper()
        if not roll:
            messagebox.showwarning('Input error', 'Enter roll to delete')
            return
        self.manager.delete(roll)
        self.clear_fields()
        self.refresh_all()

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv')
        if not path:
            return
        self.manager.export_csv(path)
        messagebox.showinfo('Exported', f'CSV exported to {path}')

    def clear_fields(self):
        self.roll_var.set('')
        self.name_var.set('')
        self.mark_var.set('')

    def on_select(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        if len(vals) >= 3:
            self.roll_var.set(vals[0])
            self.name_var.set(vals[1])
            self.mark_var.set(vals[2])

    # ---------------- Theme ----------------
    def toggle_theme(self):
        self.set_theme(not self.dark)
        # update chart styling immediately
        self.refresh_all()

    # ---------------- Charts & Table ----------------
    def refresh_all(self):
        # ensure manager data is loaded and normalized
        self.manager.load()
        # Normalize any bad entries and persist
        repaired = False
        for roll, s in list(self.manager.students.items()):
            if not isinstance(s, dict):
                try:
                    marks = float(s)
                except Exception:
                    marks = 0.0
                self.manager.students[roll] = {"name": "UNKNOWN", "marks": marks}
                repaired = True
        if repaired:
            try:
                self.manager.save()
            except Exception:
                pass

        self.refresh_table()
        self.refresh_stats()
        self.refresh_charts()

    def refresh_table(self):
        # clear existing rows
        for r in self.tree.get_children():
            self.tree.delete(r)

        # prepare dataset
        keys_to_show = self.manager.filter_search(self.search_var.get())
        sorted_data = self.manager.sorted(self.sort_mode.get())

        for roll, s in sorted_data.items():
            # guard and normalize
            if not isinstance(s, dict):
                try:
                    marks = float(s)
                except Exception:
                    marks = 0.0
                s = {"name": "UNKNOWN", "marks": marks}
                self.manager.students[roll] = s
                try:
                    self.manager.save()
                except Exception:
                    pass

            if roll in keys_to_show:
                name = s.get('name', 'UNKNOWN')
                marks = s.get('marks', 0.0)
                grade = calculate_grade(marks)
                self.tree.insert('', 'end', values=(roll, name, marks, grade))

    def refresh_stats(self):
        st = self.manager.stats()
        text = (
            f"Total Students: {st['total']}\n"
            f"Max Marks: {st['max']}\n"
            f"Min Marks: {st['min']}\n"
            f"Avg Marks: {st['avg']:.2f}\n \n"
            f"Author: Anadi Rathore"
        )
        self.stats_label.config(text=text)

    def _apply_chart_style(self, fig, ax):
        """Apply dark/light style consistently to a matplotlib Axes."""
        if self.dark:
            face = '#0f0f0f'
            ax.set_facecolor('#0b0b0b')
            fig.patch.set_facecolor('#0f0f0f')
            ax.tick_params(colors='white', which='both')
            for spine in ax.spines.values():
                spine.set_color('#444')
            ax.title.set_color('white')
            ax.xaxis.label.set_color('white')
            ax.yaxis.label.set_color('white')
        else:
            ax.set_facecolor('white')
            fig.patch.set_facecolor('white')
            ax.tick_params(colors='black', which='both')
            for spine in ax.spines.values():
                spine.set_color('#444')
            ax.title.set_color('black')
            ax.xaxis.label.set_color('black')
            ax.yaxis.label.set_color('black')

    def refresh_charts(self):
        # normalize again to be safe
        repaired = False
        for roll, s in list(self.manager.students.items()):
            if not isinstance(s, dict):
                try:
                    marks = float(s)
                except Exception:
                    marks = 0.0
                self.manager.students[roll] = {"name": "UNKNOWN", "marks": marks}
                repaired = True
        if repaired:
            try:
                self.manager.save()
            except Exception:
                pass

        data = self.manager.sorted(self.sort_mode.get())
        rolls = list(data.keys())
        marks = [data[r]['marks'] for r in rolls]
        names = [data[r]['name'] for r in rolls]

        # ---------------- LINE CHART ----------------
        self.ax_line.clear()
        self._apply_chart_style(self.fig_line, self.ax_line)
        if marks:
            x = list(range(1, len(marks) + 1))
            # choose colors friendly to dark theme
            line_color = '#4dd0e1' if self.dark else '#1f77b4'  # cyan-ish in dark, default blue in light
            marker_face = '#00bcd4' if self.dark else '#1f77b4'
            self.ax_line.plot(x, marks, marker='o', linestyle='-', linewidth=2, color=line_color, markerfacecolor=marker_face)
            self.ax_line.set_title('Line: Marks by Rank')
            self.ax_line.set_ylabel('Marks')
            self.ax_line.set_ylim(0, 100)
            self.ax_line.set_xticks(x)
            # annotate points lightly (small labels on hover only; no static labels to avoid clutter)
            # store line points for hover interactions
            self.line_points = [(x[i], marks[i], rolls[i], names[i], marks[i]) for i in range(len(x))]
        else:
            self.line_points = []
        self.canvas_line.draw()

        # ---------------- BAR CHART ----------------
        self.ax_bar.clear()
        self._apply_chart_style(self.fig_bar, self.ax_bar)
        self.bar_rects = []
        if marks:
            labels = [f"{r}\n{names[i]}" for i, r in enumerate(rolls)]
            x_pos = list(range(len(marks)))
            bar_color = '#ffb86b' if self.dark else '#ff7f0e'  # warm orange variants
            edge_color = '#ff9f40' if self.dark else '#d35400'
            bars = self.ax_bar.bar(x_pos, marks, align='center', color=bar_color, edgecolor=edge_color)
            self.ax_bar.set_xticks(x_pos)
            self.ax_bar.set_xticklabels(labels, rotation=0, ha='right')
            self.ax_bar.set_title('Bar: Marks by Roll/Name')
            self.ax_bar.set_ylabel('Marks')
            self.ax_bar.set_ylim(0, 100)

            # add data labels above bars
            for rect, val in zip(bars, marks):
                height = rect.get_height()
                # choose label color contrast
                label_color = 'white' if self.dark else 'black'
                self.ax_bar.annotate(f"{val:.0f}",
                                     xy=(rect.get_x() + rect.get_width() / 2, height),
                                     xytext=(0, 6),
                                     textcoords="offset points",
                                     ha='center', va='bottom',
                                     fontsize=9, color=label_color, fontweight='bold')
                self.bar_rects.append((rect, x_pos.pop(0) if x_pos else None))  # store rect

        # connect hover interactions
        # disconnect previous connectors if present
        try:
            if self.bar_cid:
                self.canvas_bar.mpl_disconnect(self.bar_cid)
            if self.line_cid:
                self.canvas_line.mpl_disconnect(self.line_cid)
        except Exception:
            pass

        # create or reuse hover annotation
        if self.hover_ann:
            try:
                self.hover_ann.remove()
            except Exception:
                pass
        # annotation is on the bar canvas figure
        self.hover_ann = self.ax_bar.annotate("", xy=(0, 0), xytext=(10, 10), textcoords="offset points",
                                             bbox=dict(boxstyle="round", fc="w" if not self.dark else "#222", ec="0.5"),
                                             color='black' if not self.dark else 'white',
                                             fontsize=9, visible=False)

        # handlers
        def on_move_bar(event):
            # only consider events on bar figure
            if event.inaxes != self.ax_bar:
                if self.hover_ann.get_visible():
                    self.hover_ann.set_visible(False)
                    self.canvas_bar.draw_idle()
                return
            found = False
            for rect, _ in self.bar_rects:
                contains, _ = rect.contains(event)
                if contains:
                    idx = None
                    # find index from rect's x (approx)
                    bx = rect.get_x() + rect.get_width() / 2
                    # find matching roll by bar center x (match by rect object)
                    # we can compute index by comparing with bar_rects order
                    try:
                        i = [r for r, _ in self.bar_rects].index(rect)
                    except Exception:
                        i = None
                    if i is not None and i < len(rolls):
                        roll = rolls[i]
                        name = names[i]
                        mark = marks[i]
                        txt = f"{roll} | {name}\nMarks: {mark:.1f}"
                        self.hover_ann.set_text(txt)
                        self.hover_ann.xy = (rect.get_x() + rect.get_width() / 2, rect.get_height())
                        # update annotation style depending on theme
                        self.hover_ann.get_bbox_patch().set_facecolor('#222' if self.dark else 'white')
                        self.hover_ann.get_bbox_patch().set_edgecolor('#888')
                        self.hover_ann.set_color('white' if self.dark else 'black')
                        self.hover_ann.set_visible(True)
                        self.canvas_bar.draw_idle()
                        found = True
                        break
            if not found and self.hover_ann.get_visible():
                self.hover_ann.set_visible(False)
                self.canvas_bar.draw_idle()

        def on_move_line(event):
            # only consider events on line axes
            if event.inaxes != self.ax_line:
                return
            if not self.line_points:
                return
            # find nearest point within a threshold in display coords
            trans = self.ax_line.transData
            # convert event to data coords (already)
            best = None
            best_dist = 15  # pixels threshold
            for xp, yp, roll, name, mark in self.line_points:
                xdisp, ydisp = trans.transform((xp, yp))
                dx = xdisp - event.x
                dy = ydisp - event.y
                dist = math.hypot(dx, dy)
                if dist < best_dist:
                    best_dist = dist
                    best = (xp, yp, roll, name, mark)
            # show a tooltip on the line figure if found
            # we'll reuse a small annotation drawn on the line axes
            if best:
                xp, yp, roll, name, mark = best
                # create annotation on the line axes
                # remove previous small annot if exists
                # we will create a transient annotation on ax_line
                if not hasattr(self, '_line_ann'):
                    self._line_ann = self.ax_line.annotate("", xy=(xp, yp), xytext=(10, 10), textcoords="offset points",
                                                           bbox=dict(boxstyle="round", fc="#222" if self.dark else 'white'),
                                                           color='white' if self.dark else 'black', fontsize=9, visible=True)
                self._line_ann.xy = (xp, yp)
                self._line_ann.set_text(f"{roll} | {name}\nMarks: {mark:.1f}")
                self._line_ann.set_visible(True)
                self.canvas_line.draw_idle()
            else:
                if hasattr(self, '_line_ann') and self._line_ann.get_visible():
                    self._line_ann.set_visible(False)
                    self.canvas_line.draw_idle()

        self.bar_cid = self.canvas_bar.mpl_connect('motion_notify_event', on_move_bar)
        self.line_cid = self.canvas_line.mpl_connect('motion_notify_event', on_move_line)

        self.canvas_bar.draw()
        self.canvas_line.draw()

    # ---------------- PDF Report ----------------
    def generate_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension='.pdf')
        if not path:
            return

        tmp_dir = tempfile.gettempdir()
        line_img = os.path.join(tmp_dir, 'line_chart_vityarthi.png')
        bar_img = os.path.join(tmp_dir, 'bar_chart_vityarthi.png')

        try:
            # save current figures with tight layout and transparent backgrounds if dark
            self.fig_line.savefig(line_img, bbox_inches='tight', facecolor=self.fig_line.get_facecolor())
            self.fig_bar.savefig(bar_img, bbox_inches='tight', facecolor=self.fig_bar.get_facecolor())
        except Exception:
            pass

        c = pdf_canvas.Canvas(path, pagesize=A4)
        width, height = A4

        # Draw custom heading for VIT-Yarthi
        c.setFont('Helvetica-Bold', 18)
        c.drawString(40, height - 60, 'VIT-Yarthi Student Report')

        # draw logo if available (try provided path)
        if os.path.exists(LOGO_PATH):
            try:
                img = Image.open(LOGO_PATH)
                iw, ih = img.size
                aspect = ih / iw
                w = 100
                h = int(w * aspect)
                c.drawImage(ImageReader(LOGO_PATH), width - 140, height - 80, width=w, height=h)
            except Exception:
                pass

        # stats
        st = self.manager.stats()
        c.setFont('Helvetica', 11)
        c.drawString(40, height - 100, f"Total Students: {st['total']}")
        c.drawString(200, height - 100, f"Max: {st['max']}")
        c.drawString(320, height - 100, f"Min: {st['min']}")
        c.drawString(420, height - 100, f"Avg: {st['avg']:.2f}")

        # table header
        c.setFont('Helvetica-Bold', 10)
        y = height - 130
        c.drawString(40, y, 'Rank')
        c.drawString(80, y, 'Roll')
        c.drawString(150, y, 'Name')
        c.drawString(350, y, 'Marks')
        c.drawString(420, y, 'Grade')
        y -= 14

        # table rows (defensive)
        ranked = sorted(self.manager.students.items(), key=lambda x: (x[1]['marks'] if isinstance(x[1], dict) else float(x[1])), reverse=True)
        c.setFont('Helvetica', 10)
        for i, (r, s) in enumerate(ranked, 1):
            # normalize per-row if necessary
            if not isinstance(s, dict):
                try:
                    marks_val = float(s)
                except Exception:
                    marks_val = 0.0
                s = {"name": "UNKNOWN", "marks": marks_val}
                self.manager.students[r] = s
                try:
                    self.manager.save()
                except Exception:
                    pass

            if y < 120:
                c.showPage()
                y = height - 60
            c.drawString(40, y, str(i))
            c.drawString(80, y, str(r))
            c.drawString(150, y, str(s.get('name', 'UNKNOWN')))
            c.drawString(350, y, str(s.get('marks', 0.0)))
            c.drawString(420, y, calculate_grade(s.get('marks', 0.0)))
            y -= 14

        # charts pages (line + bar)
        if os.path.exists(line_img):
            c.showPage()
            try:
                c.drawImage(ImageReader(line_img), 40, height / 2 - 20, width=width - 80, height=height / 2 - 40)
            except Exception:
                pass
        if os.path.exists(bar_img):
            c.showPage()
            try:
                c.drawImage(ImageReader(bar_img), 40, height / 2 - 20, width=width - 80, height=height / 2 - 40)
            except Exception:
                pass

        c.save()
        messagebox.showinfo('Report', f'PDF report saved to {path}')


# ---------------- Run ----------------
if __name__ == '__main__':
    root = tk.Tk()
    app = GradeApp(root)
    root.mainloop()
