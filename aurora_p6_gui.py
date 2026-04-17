"""
AURORA P6 Excel Converter - GUI
Tkinter front-end for create_template.py and aurora_p6_converter.py.
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
    APP_DIR = sys._MEIPASS          # temp folder PyInstaller unpacks to
    SCRIPTS_DIR = os.path.dirname(sys.executable)   # folder containing the .exe
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    SCRIPTS_DIR = APP_DIR


def get_python():
    """Return python executable path — works frozen or not."""
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
C_GEN_BG     = "#217346"   # green for Generate Template
C_GEN_ACT    = "#185C37"
C_CNV_BG     = "#1F4E79"   # dark blue for Convert
C_CNV_ACT    = "#163759"
C_LOG_BG     = "#1E1E1E"
C_LOG_FG     = "#D4D4D4"
C_LOG_OK     = "#4EC9B0"
C_LOG_ERR    = "#F44747"
C_LOG_INFO   = "#9CDCFE"

FONT_TITLE   = ("Segoe UI", 16, "bold")
FONT_SECTION = ("Segoe UI", 10, "bold")
FONT_BODY    = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)
FONT_INSTR   = ("Segoe UI", 9)


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AURORA P6 Excel Converter")
        self.resizable(True, True)
        self.minsize(720, 680)
        self.configure(bg=C_SECTION_BG)

        self._ref_var      = tk.StringVar()
        self._template_var = tk.StringVar()

        self._build_ui()
        self._center_window(760, 760)

    # ------------------------------------------------------------------
    # Window centering
    # ------------------------------------------------------------------
    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---- Header ----
        hdr = tk.Frame(self, bg=C_HEADER_BG, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="AURORA P6 Excel Converter",
                 font=FONT_TITLE, bg=C_HEADER_BG, fg=C_HEADER_FG).pack()
        tk.Label(hdr, text="Excel  \u2192  Primavera P6 18.8 XML",
                 font=("Segoe UI", 9), bg=C_HEADER_BG, fg="#BDD7EE").pack()

        # ---- Scrollable content area ----
        outer = tk.Frame(self, bg=C_SECTION_BG)
        outer.pack(fill="both", expand=True, padx=0, pady=0)

        canvas = tk.Canvas(outer, bg=C_SECTION_BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = tk.Frame(canvas, bg=C_SECTION_BG)
        cwin    = canvas.create_window((0, 0), window=content, anchor="nw")

        def _on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_resize(event):
            canvas.itemconfig(cwin, width=event.width)

        content.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_canvas_resize)

        # Mouse-wheel scroll
        def _on_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)

        pad = dict(padx=16, pady=8)

        # ---- Step 1 ----
        s1 = self._section(content, "Step 1 \u2014 P6 Reference File")
        s1.pack(fill="x", **pad)
        tk.Label(s1, text=(
            "Export any project from P6: File \u2192 Export \u2192 Primavera P6 XML.\n"
            "Rename the exported file and select it below."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=650).pack(
            anchor="w", padx=8, pady=(4, 6))

        row1 = tk.Frame(s1, bg=C_SECTION_BG)
        row1.pack(fill="x", padx=8, pady=(0, 6))
        self._ref_entry = tk.Entry(row1, textvariable=self._ref_var,
                                   font=FONT_BODY, width=55)
        self._ref_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self._btn(row1, "Browse\u2026", self._browse_ref,
                  bg=C_BTN_BG, abg=C_BTN_ACT).pack(side="left", padx=(6, 0))

        self._btn(s1, "Generate Template  (create_template.py)",
                  self._run_generate,
                  bg=C_GEN_BG, abg=C_GEN_ACT,
                  font=("Segoe UI", 10, "bold"), pady=6
                  ).pack(padx=8, pady=(2, 10), anchor="w")

        # ---- Step 2 ----
        s2 = self._section(content, "Step 2 \u2014 Convert Filled Template")
        s2.pack(fill="x", **pad)
        tk.Label(s2, text=(
            "Open P6_Import_Template.xlsx, fill in your data, save, then select it below."
        ), font=FONT_BODY, bg=C_SECTION_BG, justify="left", wraplength=650).pack(
            anchor="w", padx=8, pady=(4, 6))

        row2 = tk.Frame(s2, bg=C_SECTION_BG)
        row2.pack(fill="x", padx=8, pady=(0, 6))
        self._tmpl_entry = tk.Entry(row2, textvariable=self._template_var,
                                    font=FONT_BODY, width=55)
        self._tmpl_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self._btn(row2, "Browse\u2026", self._browse_template,
                  bg=C_BTN_BG, abg=C_BTN_ACT).pack(side="left", padx=(6, 0))

        self._btn(s2, "Convert to XML  (aurora_p6_converter.py)",
                  self._run_convert,
                  bg=C_CNV_BG, abg=C_CNV_ACT,
                  font=("Segoe UI", 10, "bold"), pady=6
                  ).pack(padx=8, pady=(2, 10), anchor="w")

        # ---- Step 3 — Import Instructions ----
        s3 = self._section(content, "Step 3 \u2014 Import into P6 Professional 18.8")
        s3.pack(fill="x", **pad)

        instr = (
            "Pass 1 \u2014 Create New Project\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass1.xml\n"
            "  3. Import action: Create New Project\n"
            "  4. Complete the import wizard\n"
            "  Imports: Project, WBS, Activities, Resources, ActivityCodes.\n"
            "\n"
            "Pass 2 \u2014 Add Relationships\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass2.xml\n"
            "  3. Import action: Update Existing Project\n"
            "  4. Complete the import wizard\n"
            "  Imports: Activities (matched by ObjectId) + Relationships.\n"
            "\n"
            "Pass 3 \u2014 Add Resource Assignments\n"
            "  1. File \u2192 Import \u2192 Primavera P6 XML\n"
            "  2. Select  P6_Import_pass3.xml\n"
            "  3. Import action: Update Existing Project\n"
            "  4. Complete the import wizard\n"
            "  Imports: Activities (matched by ObjectId) + ResourceAssignments.\n"
            "\n"
            "Important: All three passes must be imported in order. Do not skip Pass 1.\n"
            "Pass 3 is only required if your template has ResourceAssignment rows."
        )
        txt = tk.Text(s3, font=FONT_INSTR, bg="#EBF2FA", fg="#1A1A2E",
                      relief="flat", bd=0, height=19, wrap="word",
                      state="normal", cursor="arrow")
        txt.insert("1.0", instr)
        txt.configure(state="disabled")
        txt.pack(fill="x", padx=8, pady=(4, 10))

        # ---- Log area ----
        s4 = self._section(content, "Log")
        s4.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        log_frame = tk.Frame(s4, bg=C_LOG_BG)
        log_frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self._log = scrolledtext.ScrolledText(
            log_frame, font=FONT_MONO,
            bg=C_LOG_BG, fg=C_LOG_FG,
            insertbackground=C_LOG_FG,
            relief="flat", bd=0,
            height=10, wrap="word",
            state="disabled"
        )
        self._log.pack(fill="both", expand=True)
        self._log.tag_configure("ok",   foreground=C_LOG_OK)
        self._log.tag_configure("err",  foreground=C_LOG_ERR)
        self._log.tag_configure("info", foreground=C_LOG_INFO)

        # Clear log button
        self._btn(s4, "Clear Log", self._clear_log,
                  bg="#555", abg="#333", font=FONT_BODY
                  ).pack(anchor="e", padx=8, pady=(0, 6))

    # ------------------------------------------------------------------
    # Helper: labelled section frame
    # ------------------------------------------------------------------
    def _section(self, parent, title):
        lf = tk.LabelFrame(parent, text=f"  {title}  ",
                           font=FONT_SECTION,
                           bg=C_SECTION_BG, fg="#1F4E79",
                           relief="groove", bd=2, labelanchor="nw")
        return lf

    # ------------------------------------------------------------------
    # Helper: styled button
    # ------------------------------------------------------------------
    def _btn(self, parent, text, cmd, bg=C_BTN_BG, abg=C_BTN_ACT,
             font=FONT_BODY, pady=4):
        b = tk.Button(parent, text=text, command=cmd,
                      font=font,
                      bg=bg, fg=C_BTN_FG,
                      activebackground=abg, activeforeground=C_BTN_FG,
                      relief="flat", padx=12, pady=pady,
                      cursor="hand2")
        b.bind("<Enter>", lambda e: b.configure(bg=abg))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    # ------------------------------------------------------------------
    # File browsing
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
    # Ensure p6_reference.xml is next to the scripts
    # ------------------------------------------------------------------
    def _prepare_ref(self, dest_dir):
        """Copy selected reference file to dest_dir as p6_reference.xml.
        Returns True on success, False on error."""
        ref_src = self._ref_var.get().strip()
        if not ref_src:
            messagebox.showerror("Missing Reference File",
                                 "Please select a P6 reference XML file (Step 1).")
            return False
        if not os.path.isfile(ref_src):
            messagebox.showerror("File Not Found",
                                 f"Reference file not found:\n{ref_src}")
            return False
        dest = os.path.join(dest_dir, "p6_reference.xml")
        if os.path.abspath(ref_src) != os.path.abspath(dest):
            try:
                shutil.copy2(ref_src, dest)
                self._log_line(f"Copied reference file -> {dest}", "info")
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
                    text=True,
                    encoding="utf-8", errors="replace"
                )
                for line in proc.stdout:
                    tag = "err" if any(w in line.lower()
                                       for w in ("error", "traceback", "exception"))  \
                          else None
                    self.after(0, self._log_line, line, tag)
                proc.wait()
                self.after(0, on_done, proc.returncode)
            except Exception as exc:
                self.after(0, self._log_line, f"ERROR: {exc}\n", "err")
                self.after(0, on_done, -1)
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Step 1 — Generate Template
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

        # Work from SCRIPTS_DIR so that outputs land next to the converter
        work_dir = SCRIPTS_DIR

        if not self._prepare_ref(work_dir):
            return

        self._log_line("--- Generate Template ---", "info")

        def done(rc):
            if rc == 0:
                out = os.path.join(work_dir, "P6_Import_Template.xlsx")
                self._log_line("Template generated successfully.", "ok")
                messagebox.showinfo(
                    "Template Generated",
                    f"P6_Import_Template.xlsx created in:\n{work_dir}\n\n"
                    "Open it, fill in your project data, then use Step 2 to convert."
                )
            else:
                self._log_line("Template generation failed. See log above.", "err")
                messagebox.showerror("Error",
                                     "create_template.py exited with an error.\n"
                                     "See the Log area for details.")

        self._run_in_thread([py, script], work_dir, done)

    # ------------------------------------------------------------------
    # Step 2 — Convert Template to XML
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
            messagebox.showerror("File Not Found",
                                 f"Template not found:\n{template}")
            return

        script = os.path.join(APP_DIR, "aurora_p6_converter.py")
        if not os.path.isfile(script):
            messagebox.showerror("Script Not Found",
                                 f"aurora_p6_converter.py not found in:\n{APP_DIR}")
            return

        # Reference file must live next to the template
        tmpl_dir = os.path.dirname(os.path.abspath(template))
        if not self._prepare_ref(tmpl_dir):
            return

        self._log_line("--- Convert to XML ---", "info")

        def done(rc):
            if rc == 0:
                self._log_line("Conversion complete.", "ok")
                messagebox.showinfo(
                    "Conversion Complete",
                    f"Three XML files created in:\n{tmpl_dir}\n\n"
                    "  P6_Import_pass1.xml  \u2014  Create New Project\n"
                    "  P6_Import_pass2.xml  \u2014  Update Existing Project (Relationships)\n"
                    "  P6_Import_pass3.xml  \u2014  Update Existing Project (ResourceAssignments)\n\n"
                    "See Step 3 for import instructions."
                )
            else:
                self._log_line("Conversion failed. See log above.", "err")
                messagebox.showerror("Error",
                                     "aurora_p6_converter.py exited with an error.\n"
                                     "See the Log area for details.")

        self._run_in_thread([py, script, template], tmpl_dir, done)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
