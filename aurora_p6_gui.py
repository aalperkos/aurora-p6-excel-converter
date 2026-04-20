"""
AURORA P6 Converter - GUI
Tkinter front-end for:
  create_template.py        (Excel template generator)
  aurora_p6_converter.py    (Excel → P6 XML)
  msp_to_p6_converter.py    (MS Project XML → P6 XML)
"""

import os
import sys
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

# ---------------------------------------------------------------------------
# Path resolution — works both frozen (PyInstaller) and from source
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    APP_DIR    = sys._MEIPASS
    SCRIPTS_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR    = os.path.dirname(os.path.abspath(__file__))
    SCRIPTS_DIR = APP_DIR


def get_python():
    if getattr(sys, "frozen", False):
        py = shutil.which("py") or shutil.which("python3") or shutil.which("python")
        if py:
            return py
        messagebox.showerror(
            "Python Not Found",
            "Could not locate Python. Install Python 3.9+ and ensure it is on PATH."
        )
        return None
    return sys.executable


# ---------------------------------------------------------------------------
# Colours / fonts
# ---------------------------------------------------------------------------
C_HEADER_BG  = "#1F4E79"
C_HEADER_FG  = "#FFFFFF"
C_SECTION_BG = "#F0F4F8"
C_BTN_BG     = "#2E75B6"
C_BTN_FG     = "#FFFFFF"
C_BTN_ACT    = "#1F5B8E"
C_GEN_BG     = "#217346"
C_GEN_ACT    = "#185C37"
C_CNV_BG     = "#1F4E79"
C_CNV_ACT    = "#163759"
C_MSP_BG     = "#5B4A9E"
C_MSP_ACT    = "#42357A"
C_RBD_BG     = "#B85C00"
C_RBD_ACT    = "#8A4400"
C_LOG_BG     = "#1E1E1E"
C_LOG_FG     = "#D4D4D4"
C_LOG_OK     = "#4EC9B0"
C_LOG_ERR    = "#F44747"
C_LOG_INFO   = "#9CDCFE"
C_LOG_WARN   = "#CE9178"
C_DISABLED   = "#AAAAAA"

FONT_TITLE   = ("Segoe UI", 16, "bold")
FONT_SECTION = ("Segoe UI", 10, "bold")
FONT_BODY    = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)
FONT_INSTR   = ("Segoe UI", 9)
FONT_PATH    = ("Consolas", 8)


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AURORA P6 Converter")
        self.resizable(True, True)
        self.minsize(760, 720)
        self.configure(bg=C_SECTION_BG)

        # Excel tab state
        self._ref_var      = tk.StringVar()
        self._template_var = tk.StringVar()

        # MSP tab state
        self._msp_var      = tk.StringVar()
        self._msp_ref_var  = tk.StringVar()
        self._msp_path     = None   # resolved after convert
        self._msp_out_dir  = None
        self._msp_base     = None
        self._rebind_btn   = None   # enabled after convert

        self._build_ui()
        self._center_window(820, 860)

    # ------------------------------------------------------------------
    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---- Header ----
        hdr = tk.Frame(self, bg=C_HEADER_BG, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="AURORA P6 Converter",
                 font=FONT_TITLE, bg=C_HEADER_BG, fg=C_HEADER_FG).pack()
        tk.Label(hdr,
                 text="Excel \u2192 P6 XML  |  MS Project XML \u2192 P6 XML",
                 font=("Segoe UI", 9), bg=C_HEADER_BG, fg="#BDD7EE").pack()

        # ---- Notebook ----
        style = ttk.Style(self)
        style.configure("TNotebook",        background=C_SECTION_BG)
        style.configure("TNotebook.Tab",    font=("Segoe UI", 10, "bold"),
                                            padding=[14, 6])
        style.configure("TFrame",           background=C_SECTION_BG)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        tab1 = tk.Frame(nb, bg=C_SECTION_BG)
        nb.add(tab1, text="  Excel \u2192 P6  ")
        self._build_excel_tab(tab1)

        tab2 = tk.Frame(nb, bg=C_SECTION_BG)
        nb.add(tab2, text="  MS Project \u2192 P6  ")
        self._build_msp_tab(tab2)

        # ---- Shared log ----
        log_outer = tk.Frame(self, bg=C_SECTION_BG)
        log_outer.pack(fill="x", padx=12, pady=(4, 8))

        log_lf = tk.LabelFrame(log_outer, text="  Log  ",
                                font=FONT_SECTION, bg=C_SECTION_BG,
                                fg="#1F4E79", relief="groove", bd=2,
                                labelanchor="nw")
        log_lf.pack(fill="x")

        log_frame = tk.Frame(log_lf, bg=C_LOG_BG)
        log_frame.pack(fill="x", padx=8, pady=(4, 4))

        self._log = scrolledtext.ScrolledText(
            log_frame, font=FONT_MONO,
            bg=C_LOG_BG, fg=C_LOG_FG,
            insertbackground=C_LOG_FG,
            relief="flat", bd=0,
            height=8, wrap="word",
            state="disabled"
        )
        self._log.pack(fill="x")
        self._log.tag_configure("ok",   foreground=C_LOG_OK)
        self._log.tag_configure("err",  foreground=C_LOG_ERR)
        self._log.tag_configure("info", foreground=C_LOG_INFO)
        self._log.tag_configure("warn", foreground=C_LOG_WARN)

        self._btn(log_lf, "Clear Log", self._clear_log,
                  bg="#555", abg="#333", font=FONT_BODY
                  ).pack(anchor="e", padx=8, pady=(0, 6))

    # ------------------------------------------------------------------
    # Helper: scrollable frame inside a tab
    # ------------------------------------------------------------------
    def _make_scrollable(self, parent):
        canvas = tk.Canvas(parent, bg=C_SECTION_BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = tk.Frame(canvas, bg=C_SECTION_BG)
        cwin    = canvas.create_window((0, 0), window=content, anchor="nw")

        content.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(cwin, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        return canvas, content

    # ------------------------------------------------------------------
    # Helper: section LabelFrame
    # ------------------------------------------------------------------
    def _section(self, parent, title):
        return tk.LabelFrame(parent, text=f"  {title}  ",
                             font=FONT_SECTION,
                             bg=C_SECTION_BG, fg="#1F4E79",
                             relief="groove", bd=2, labelanchor="nw")

    # ------------------------------------------------------------------
    # Helper: styled button
    # ------------------------------------------------------------------
    def _btn(self, parent, text, cmd, bg=C_BTN_BG, abg=C_BTN_ACT,
             font=FONT_BODY, pady=4):
        b = tk.Button(parent, text=text, command=cmd,
                      font=font, bg=bg, fg=C_BTN_FG,
                      activebackground=abg, activeforeground=C_BTN_FG,
                      relief="flat", padx=12, pady=pady, cursor="hand2")
        b.bind("<Enter>", lambda e: b.configure(bg=abg))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    # ------------------------------------------------------------------
    # Helper: file row (Entry + Browse button)
    # ------------------------------------------------------------------
    def _file_row(self, parent, var, browse_cmd):
        row = tk.Frame(parent, bg=C_SECTION_BG)
        row.pack(fill="x", padx=8, pady=(0, 6))
        tk.Entry(row, textvariable=var, font=FONT_BODY, width=58
                 ).pack(side="left", fill="x", expand=True, ipady=3)
        self._btn(row, "Browse\u2026", browse_cmd,
                  bg=C_BTN_BG, abg=C_BTN_ACT).pack(side="left", padx=(6, 0))

    # ======================================================================
    # EXCEL TAB
    # ======================================================================
    def _build_excel_tab(self, parent):
        _, content = self._make_scrollable(parent)
        pad = dict(padx=16, pady=8)

        # Step 1
        s1 = self._section(content, "Step 1 \u2014 P6 Reference File")
        s1.pack(fill="x", **pad)
        tk.Label(s1, text=(
            "Export any project from P6: File \u2192 Export \u2192 Primavera P6 XML.\n"
            "Rename the exported file and select it below."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=680
        ).pack(anchor="w", padx=8, pady=(4, 6))

        self._file_row(s1, self._ref_var, self._browse_ref)

        self._btn(s1, "Generate Template  (create_template.py)",
                  self._run_generate,
                  bg=C_GEN_BG, abg=C_GEN_ACT,
                  font=("Segoe UI", 10, "bold"), pady=6
                  ).pack(padx=8, pady=(2, 10), anchor="w")

        # Step 2
        s2 = self._section(content, "Step 2 \u2014 Convert Filled Template")
        s2.pack(fill="x", **pad)
        tk.Label(s2, text=(
            "Open P6_Import_Template.xlsx, fill in your data, save, then select it below."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=680
        ).pack(anchor="w", padx=8, pady=(4, 6))

        self._file_row(s2, self._template_var, self._browse_template)

        self._btn(s2, "Convert to XML  (aurora_p6_converter.py)",
                  self._run_convert,
                  bg=C_CNV_BG, abg=C_CNV_ACT,
                  font=("Segoe UI", 10, "bold"), pady=6
                  ).pack(padx=8, pady=(2, 10), anchor="w")

        # Step 3 — import instructions
        s3 = self._section(content, "Step 3 \u2014 Import into P6 Professional 18.8")
        s3.pack(fill="x", **pad)
        self._instr_text(s3, (
            "Pass 1 \u2014 Create New Project\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass1.xml\n"
            "  3. Import action: Create New Project\n"
            "  Imports: Project, WBS, Activities, Resources, ActivityCodes.\n"
            "\n"
            "Pass 2 \u2014 Add Relationships\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass2.xml\n"
            "  3. Import action: Update Existing Project\n"
            "  Imports: Activities (matched by ObjectId) + Relationships.\n"
            "\n"
            "Pass 3 \u2014 Add Resource Assignments\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass3.xml\n"
            "  3. Import action: Update Existing Project\n"
            "  Imports: Activities (matched by ObjectId) + ResourceAssignments.\n"
            "\n"
            "Important: All three passes must be imported in order. Do not skip Pass 1.\n"
            "Pass 3 is only required if your template has ResourceAssignment rows."
        ), height=19)

    # ======================================================================
    # MSP TAB
    # ======================================================================
    def _build_msp_tab(self, parent):
        _, content = self._make_scrollable(parent)
        pad = dict(padx=16, pady=8)

        # ---- Section 1: MSP XML ----
        s1 = self._section(content, "Step 1 \u2014 MS Project XML File")
        s1.pack(fill="x", **pad)
        tk.Label(s1, text=(
            "Export your schedule from Microsoft Project:\n"
            "  File \u2192 Save As \u2192 XML Format (*.xml)"
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=680
        ).pack(anchor="w", padx=8, pady=(4, 6))

        self._file_row(s1, self._msp_var, self._browse_msp)

        # ---- Section 2: Reference file ----
        s2 = self._section(content, "Step 2 \u2014 P6 Reference File")
        s2.pack(fill="x", **pad)
        tk.Label(s2, text=(
            "Required: export any project from P6 (File \u2192 Export \u2192 Primavera P6 XML).\n"
            "Provides Calendar, OBS, and EPS anchor for the new project.\n"
            "Auto-detect searches the MS Project XML folder and the script folder."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=680
        ).pack(anchor="w", padx=8, pady=(4, 4))

        ref_row = tk.Frame(s2, bg=C_SECTION_BG)
        ref_row.pack(fill="x", padx=8, pady=(0, 8))
        tk.Entry(ref_row, textvariable=self._msp_ref_var,
                 font=FONT_BODY, width=48
                 ).pack(side="left", fill="x", expand=True, ipady=3)
        self._btn(ref_row, "Browse\u2026", self._browse_msp_ref,
                  bg=C_BTN_BG, abg=C_BTN_ACT).pack(side="left", padx=(6, 0))
        self._btn(ref_row, "Auto-detect", self._autodetect_msp_ref,
                  bg="#555", abg="#333").pack(side="left", padx=(6, 0))

        # ---- Section 3: Convert ----
        s3 = self._section(content, "Step 3 \u2014 Convert to P6 XML")
        s3.pack(fill="x", **pad)

        self._btn(s3, "Convert to P6 XML  (msp_to_p6_converter.py)",
                  self._run_msp_convert,
                  bg=C_MSP_BG, abg=C_MSP_ACT,
                  font=("Segoe UI", 10, "bold"), pady=6
                  ).pack(padx=8, pady=(8, 6), anchor="w")

        # Output file paths (updated after conversion)
        out_frame = tk.Frame(s3, bg="#E8EEF4", relief="flat", bd=0)
        out_frame.pack(fill="x", padx=8, pady=(0, 10))

        tk.Label(out_frame, text="Output files:", font=FONT_SECTION,
                 bg="#E8EEF4", fg="#1F4E79").pack(anchor="w", padx=8, pady=(6, 2))

        self._msp_pass1_var = tk.StringVar(value="  (not yet generated)")
        self._msp_pass2_var = tk.StringVar(value="  (not yet generated)")
        self._msp_pass3_var = tk.StringVar(value="  (not yet generated)")

        for label, var in [("Pass 1:", self._msp_pass1_var),
                           ("Pass 2:", self._msp_pass2_var),
                           ("Pass 3:", self._msp_pass3_var)]:
            row = tk.Frame(out_frame, bg="#E8EEF4")
            row.pack(fill="x", padx=8, pady=1)
            tk.Label(row, text=label, font=FONT_SECTION, bg="#E8EEF4",
                     fg="#555555", width=8, anchor="w").pack(side="left")
            tk.Label(row, textvariable=var, font=FONT_PATH, bg="#E8EEF4",
                     fg="#1A1A2E", anchor="w").pack(side="left", fill="x")

        tk.Frame(out_frame, bg="#E8EEF4", height=6).pack()

        # ---- Section 4: Rebind ----
        s4 = self._section(content, "Step 4 \u2014 Rebind Pass 2 & 3  (after Pass 1 import)")
        s4.pack(fill="x", **pad)
        tk.Label(s4, text=(
            "After importing Pass 1 into P6, P6 assigns its own internal ObjectIds.\n"
            "Click Rebind to query the P6 SQL Server database and rewrite Pass 2 & 3\n"
            "with the real ObjectIds so relationships and resources import correctly.\n"
            "Requires SQL Server (PBZA) running on localhost."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=680
        ).pack(anchor="w", padx=8, pady=(4, 6))

        self._rebind_btn = self._btn(
            s4, "Rebind Pass2 & Pass3",
            self._run_msp_rebind,
            bg=C_RBD_BG, abg=C_RBD_ACT,
            font=("Segoe UI", 10, "bold"), pady=6
        )
        self._rebind_btn.pack(padx=8, pady=(0, 10), anchor="w")
        self._rebind_btn.configure(state="disabled",
                                   bg=C_DISABLED, activebackground=C_DISABLED)
        self._rebind_btn.unbind("<Enter>")
        self._rebind_btn.unbind("<Leave>")

        # ---- Section 5: Import instructions ----
        s5 = self._section(content, "Step 5 \u2014 Import into P6 Professional 18.8")
        s5.pack(fill="x", **pad)
        self._instr_text(s5, (
            "Pass 1 \u2014 Create New Project\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  {base}_P6_pass1.xml\n"
            "  3. Import action: Create New Project\n"
            "  Imports: Project, WBS, Activities, Resources.\n"
            "\n"
            "After Pass 1 import \u2014 click 'Rebind Pass2 & Pass3' (Step 4 above)\n"
            "  This queries PBZA SQL Server for real ObjectIds assigned by P6.\n"
            "  Pass 2 and Pass 3 files are rewritten with correct ObjectIds.\n"
            "\n"
            "Pass 2 \u2014 Add Relationships\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  {base}_P6_pass2.xml  (after Rebind)\n"
            "  3. Import action: Update Existing Project\n"
            "  Imports: Activities + Relationships.\n"
            "\n"
            "Pass 3 \u2014 Add Resource Assignments\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  {base}_P6_pass3.xml  (after Rebind)\n"
            "  3. Import action: Update Existing Project\n"
            "  Imports: Activities + ResourceAssignments.\n"
            "\n"
            "Important: Rebind must run after every Pass 1 import."
        ), height=22)

    # ------------------------------------------------------------------
    # Helper: read-only text widget for instructions
    # ------------------------------------------------------------------
    def _instr_text(self, parent, text, height=12):
        txt = tk.Text(parent, font=FONT_INSTR, bg="#EBF2FA", fg="#1A1A2E",
                      relief="flat", bd=0, height=height, wrap="word",
                      state="normal", cursor="arrow")
        txt.insert("1.0", text)
        txt.configure(state="disabled")
        txt.pack(fill="x", padx=8, pady=(4, 10))

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------
    def _log_write(self, text, tag=None):
        self._log.configure(state="normal")
        if tag:
            self._log.insert("end", text, tag)
        else:
            self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_line(self, text, tag=None):
        self._log_write(text.rstrip("\n") + "\n", tag)

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ------------------------------------------------------------------
    # Copy p6_reference.xml to dest_dir
    # ------------------------------------------------------------------
    def _prepare_ref(self, ref_src, dest_dir, label="reference file"):
        ref_src = (ref_src or "").strip()
        if not ref_src:
            messagebox.showerror("Missing Reference File",
                                 f"Please select a P6 reference XML file ({label}).")
            return False
        if not os.path.isfile(ref_src):
            messagebox.showerror("File Not Found",
                                 f"Reference file not found:\n{ref_src}")
            return False
        dest = os.path.join(dest_dir, "p6_reference.xml")
        if os.path.abspath(ref_src) != os.path.abspath(dest):
            try:
                shutil.copy2(ref_src, dest)
                self._log_line(f"Copied reference -> {dest}", "info")
            except Exception as exc:
                messagebox.showerror("Copy Failed", str(exc))
                return False
        return True

    # ------------------------------------------------------------------
    # Run subprocess in a thread, stream output to log
    # ------------------------------------------------------------------
    def _run_in_thread(self, cmd, cwd, on_done):
        def worker():
            self._log_line(f"Running: {' '.join(cmd)}", "info")
            try:
                proc = subprocess.Popen(
                    cmd, cwd=cwd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace"
                )
                for line in proc.stdout:
                    tag = ("err"  if any(w in line.lower()
                                         for w in ("error", "traceback", "exception"))
                           else "warn" if "warn" in line.lower()
                           else None)
                    self.after(0, self._log_line, line, tag)
                proc.wait()
                self.after(0, on_done, proc.returncode)
            except Exception as exc:
                self.after(0, self._log_line, f"ERROR: {exc}\n", "err")
                self.after(0, on_done, -1)
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Excel tab — file browsing
    # ------------------------------------------------------------------
    def _browse_ref(self):
        path = filedialog.askopenfilename(
            title="Select P6 Reference XML",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if path:
            self._ref_var.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select filled P6_Import_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self._template_var.set(path)

    # ------------------------------------------------------------------
    # Excel tab — Generate Template
    # ------------------------------------------------------------------
    def _run_generate(self):
        py = get_python()
        if not py:
            return
        script = os.path.join(APP_DIR, "create_template.py")
        if not os.path.isfile(script):
            messagebox.showerror("Script Not Found",
                                 f"create_template.py not found in:\n{APP_DIR}")
            return
        work_dir = SCRIPTS_DIR
        if not self._prepare_ref(self._ref_var.get(), work_dir, "Step 1"):
            return
        self._log_line("--- Generate Template ---", "info")

        def done(rc):
            if rc == 0:
                self._log_line("Template generated successfully.", "ok")
                messagebox.showinfo(
                    "Template Generated",
                    f"P6_Import_Template.xlsx created in:\n{work_dir}\n\n"
                    "Open it, fill in your project data, then use Step 2 to convert."
                )
            else:
                self._log_line("Template generation failed. See log.", "err")
                messagebox.showerror("Error",
                                     "create_template.py exited with an error.\n"
                                     "See the Log area for details.")

        self._run_in_thread([py, script], work_dir, done)

    # ------------------------------------------------------------------
    # Excel tab — Convert Template to XML
    # ------------------------------------------------------------------
    def _run_convert(self):
        py = get_python()
        if not py:
            return
        template = self._template_var.get().strip()
        if not template:
            messagebox.showerror("Missing Template",
                                 "Please select P6_Import_Template.xlsx (Step 2).")
            return
        if not os.path.isfile(template):
            messagebox.showerror("File Not Found", f"Template not found:\n{template}")
            return
        script = os.path.join(APP_DIR, "aurora_p6_converter.py")
        if not os.path.isfile(script):
            messagebox.showerror("Script Not Found",
                                 f"aurora_p6_converter.py not found in:\n{APP_DIR}")
            return
        tmpl_dir = os.path.dirname(os.path.abspath(template))
        if not self._prepare_ref(self._ref_var.get(), tmpl_dir, "Step 1"):
            return
        self._log_line("--- Excel Convert to XML ---", "info")

        def done(rc):
            if rc == 0:
                self._log_line("Conversion complete.", "ok")
                messagebox.showinfo(
                    "Conversion Complete",
                    f"Three XML files created in:\n{tmpl_dir}\n\n"
                    "  P6_Import_pass1.xml  \u2014  Create New Project\n"
                    "  P6_Import_pass2.xml  \u2014  Update Existing (Relationships)\n"
                    "  P6_Import_pass3.xml  \u2014  Update Existing (ResourceAssignments)\n\n"
                    "See Step 3 for import instructions."
                )
            else:
                self._log_line("Conversion failed. See log.", "err")
                messagebox.showerror("Error",
                                     "aurora_p6_converter.py exited with an error.\n"
                                     "See the Log area for details.")

        self._run_in_thread([py, script, template], tmpl_dir, done)

    # ------------------------------------------------------------------
    # MSP tab — file browsing
    # ------------------------------------------------------------------
    def _browse_msp(self):
        path = filedialog.askopenfilename(
            title="Select MS Project XML file",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if path:
            self._msp_var.set(path)
            self._autodetect_msp_ref()

    def _browse_msp_ref(self):
        path = filedialog.askopenfilename(
            title="Select P6 Reference XML",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if path:
            self._msp_ref_var.set(path)

    def _autodetect_msp_ref(self):
        """Look for p6_reference.xml near the MSP file or in script dir."""
        candidates = []
        msp = self._msp_var.get().strip()
        if msp:
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(msp)),
                                           "p6_reference.xml"))
        candidates.append(os.path.join(SCRIPTS_DIR, "p6_reference.xml"))
        candidates.append(os.path.join(APP_DIR, "p6_reference.xml"))

        for c in candidates:
            if os.path.isfile(c):
                self._msp_ref_var.set(c)
                self._log_line(f"Auto-detected p6_reference.xml: {c}", "info")
                return
        self._log_line("Auto-detect: p6_reference.xml not found. Browse manually.", "warn")

    # ------------------------------------------------------------------
    # MSP tab — Convert
    # ------------------------------------------------------------------
    def _run_msp_convert(self):
        py = get_python()
        if not py:
            return

        msp = self._msp_var.get().strip()
        if not msp:
            messagebox.showerror("Missing File", "Please select an MS Project XML file (Step 1).")
            return
        if not os.path.isfile(msp):
            messagebox.showerror("File Not Found", f"MS Project XML not found:\n{msp}")
            return

        script = os.path.join(APP_DIR, "msp_to_p6_converter.py")
        if not os.path.isfile(script):
            messagebox.showerror("Script Not Found",
                                 f"msp_to_p6_converter.py not found in:\n{APP_DIR}")
            return

        # p6_reference.xml must be in APP_DIR (script_dir) for the converter
        if not self._prepare_ref(self._msp_ref_var.get(), APP_DIR, "Step 2"):
            return

        out_dir = os.path.dirname(os.path.abspath(msp))
        base    = os.path.splitext(os.path.basename(msp))[0]

        self._log_line("--- MSP Convert to P6 XML ---", "info")

        def done(rc):
            if rc == 0:
                p1 = os.path.join(out_dir, f"{base}_P6_pass1.xml")
                p2 = os.path.join(out_dir, f"{base}_P6_pass2.xml")
                p3 = os.path.join(out_dir, f"{base}_P6_pass3.xml")
                self._msp_pass1_var.set(p1)
                self._msp_pass2_var.set(p2)
                self._msp_pass3_var.set(p3)
                self._msp_path    = msp
                self._msp_out_dir = out_dir
                self._msp_base    = base

                # Enable rebind button
                self._rebind_btn.configure(state="normal",
                                           bg=C_RBD_BG,
                                           activebackground=C_RBD_ACT)
                self._rebind_btn.bind("<Enter>",
                                      lambda e: self._rebind_btn.configure(bg=C_RBD_ACT))
                self._rebind_btn.bind("<Leave>",
                                      lambda e: self._rebind_btn.configure(bg=C_RBD_BG))

                self._log_line("Conversion complete.", "ok")
                messagebox.showinfo(
                    "Conversion Complete",
                    f"Three XML files created in:\n{out_dir}\n\n"
                    f"  {base}_P6_pass1.xml  \u2014  Create New Project\n"
                    f"  {base}_P6_pass2.xml  \u2014  Update Existing (Relationships)\n"
                    f"  {base}_P6_pass3.xml  \u2014  Update Existing (ResourceAssignments)\n\n"
                    "Next: Import Pass 1 into P6, then click Rebind (Step 4)."
                )
            else:
                self._log_line("Conversion failed. See log.", "err")
                messagebox.showerror("Error",
                                     "msp_to_p6_converter.py exited with an error.\n"
                                     "See the Log area for details.")

        self._run_in_thread([py, script, msp, out_dir], APP_DIR, done)

    # ------------------------------------------------------------------
    # MSP tab — Rebind
    # ------------------------------------------------------------------
    def _run_msp_rebind(self):
        if not self._msp_path or not self._msp_out_dir:
            messagebox.showerror("Not Ready",
                                 "Run Convert first (Step 3), then import Pass 1 into P6.")
            return

        py = get_python()
        if not py:
            return

        script = os.path.join(APP_DIR, "msp_to_p6_converter.py")
        if not os.path.isfile(script):
            messagebox.showerror("Script Not Found",
                                 f"msp_to_p6_converter.py not found in:\n{APP_DIR}")
            return

        # Ensure reference is still in place
        if not self._prepare_ref(self._msp_ref_var.get(), APP_DIR, "Step 2"):
            return

        self._log_line("--- MSP Rebind Pass2 & Pass3 ---", "info")

        msp_path = self._msp_path
        out_dir  = self._msp_out_dir
        base     = self._msp_base

        def done(rc):
            if rc == 0:
                p2 = os.path.join(out_dir, f"{base}_P6_pass2.xml")
                p3 = os.path.join(out_dir, f"{base}_P6_pass3.xml")
                self._msp_pass2_var.set(f"{p2}  [rebound]")
                self._msp_pass3_var.set(f"{p3}  [rebound]")
                self._log_line("Rebind complete. Import Pass 2 then Pass 3.", "ok")
                messagebox.showinfo(
                    "Rebind Complete",
                    f"Pass 2 and Pass 3 rewritten with real P6 ObjectIds.\n\n"
                    f"  {base}_P6_pass2.xml  \u2014  Update Existing Project (Relationships)\n"
                    f"  {base}_P6_pass3.xml  \u2014  Update Existing Project (ResourceAssignments)\n\n"
                    "Import Pass 2, then Pass 3 into P6."
                )
            else:
                self._log_line("Rebind failed. See log.", "err")
                messagebox.showerror("Error",
                                     "Rebind failed.\n"
                                     "Ensure P6 SQL Server (PBZA) is running and Pass 1 was imported.\n"
                                     "See the Log area for details.")

        self._run_in_thread(
            [py, script, msp_path, "--rebind", out_dir],
            APP_DIR, done
        )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
