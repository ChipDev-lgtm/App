#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
CV Lense – Offline Resume Screener (Windows-ready, same design)
- Keeps your CustomTkinter dark UI exactly as-is.
- Fixes: __init__, __file__, resource_path, open-with-default on Windows.
- Multi-file: add many PDFs/ZIPs at once (dialog or drag & drop). Ranks all.
- Safe imports, PDF parsing fallbacks, CSV/JSON exports, preview formatting.
"""

from __future__ import annotations
import os, re, sys, csv, json, time, uuid, math, shutil, zipfile, tempfile, hashlib, subprocess, platform, textwrap
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional

# ---------------- Basics / Tk ----------------
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

def resource_path(p: str) -> str:
    """
    Resolve resource path both in dev and when bundled (e.g., PyInstaller --onefile).
    """
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, p)

# ---- Optional PDF text extractor ----
try:
    # Some users install an alias; try this name first
    from pdfminer_high_level import extract_text as pdfminer_extract_text  # type: ignore
    HAS_PDFMINER = True
except Exception:
    try:
        from pdfminer.high_level import extract_text as pdfminer_extract_text  # type: ignore
        HAS_PDFMINER = True
    except Exception:
        pdfminer_extract_text = None
        HAS_PDFMINER = False

# ---- Optional PDF export (ReportLab) ----
try:
    from reportlab.lib.pagesizes import A4  # type: ignore
    from reportlab.pdfgen import canvas    # type: ignore
    HAS_REPORTLAB = True
except Exception:
    HAS_REPORTLAB = False

# ---- CustomTkinter UI ----
try:
    import customtkinter as ctk
except Exception as e:
    raise SystemExit("customtkinter is required. Install: pip install customtkinter") from e

# ---- Optional Drag & Drop ----
HAS_DND = False
try:
    from TkinterDnD2 import DND_FILES, TkinterDnD  # type: ignore
    # forcing off for Windows stability; enable if your environment supports it reliably
    HAS_DND = False
except Exception:
    HAS_DND = False

APP_NAME = "CV Lense – Offline Resume Screener"
APP_ID = "cvlense.v1"
TRIAL_DAYS = 7
SIDEBAR_W = 380

# ---------------- Licensing -----------------
class LicenseManager:
    def __init__(self, app_id: str = APP_ID):
        self.license_path = os.path.join(os.path.expanduser("~"), f".{self._safe(app_id)}.license.json")
        self.machine_fingerprint = self._machine_fingerprint()
        self.license = self._load()

    def _safe(self, s: str) -> str:
        return re.sub(r"[^a-zA-Z0-9_.-]", "", s)

    def _machine_fingerprint(self) -> str:
        mac = uuid.getnode()
        data = f"{mac}-{sys.platform}"
        return hashlib.sha256(data.encode()).hexdigest()[:32]

    def _load(self) -> Dict:
        if os.path.exists(self.license_path):
            try:
                with open(self.license_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        now = int(time.time())
        return {"fingerprint": self.machine_fingerprint, "plan": "trial", "started": now, "last_check": now}

    def save(self):
        self.license["fingerprint"] = self.machine_fingerprint
        with open(self.license_path, "w", encoding="utf-8") as f:
            json.dump(self.license, f, indent=2)

    def set_plan(self, plan: str):
        assert plan in ("lifetime", "monthly", "trial")
        self.license["plan"] = plan
        if plan == "monthly":
            self.license["paid_until"] = int(time.time()) + 30*24*3600
        elif plan == "lifetime":
            self.license["paid_until"] = 2_147_483_600
        self.save()

    def ok(self) -> Tuple[bool, str]:
        if self.license.get("fingerprint") != self.machine_fingerprint:
            return False, "License fingerprint mismatch."
        plan = self.license.get("plan", "trial")
        now = int(time.time())
        if plan == "trial":
            started = self.license.get("started", now)
            days_used = (now - started) / 86400
            if days_used <= TRIAL_DAYS:
                return True, f"Trial: {TRIAL_DAYS - int(days_used)} day(s) left"
            return False, "Trial expired. Choose Monthly ($10) or Lifetime ($99)."
        paid_until = self.license.get("paid_until", 0)
        if now <= paid_until:
            return True, plan.capitalize()
        return False, "Subscription expired. Renew Monthly or upgrade to Lifetime."

# ---------------- Data Structures -----------------
@dataclass
class Candidate:
    name: str
    email: Optional[str]
    phone: Optional[str]
    path: str
    text: str
    score: float = 0.0
    matched: Dict[str, float] = field(default_factory=dict)
    summary: str = ""

# ---------------- CV Processing -----------------
class CVProcessor:
    EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
    PHONE_RE = re.compile(r"(?:(?:\+\d{1,3}[- ]?)?\d{3}[- ]?\d{3}[- ]?\d{4})")
    NAME_LINE_RE = re.compile(r"^[A-Z][A-Za-z'’-]+(?: [A-Z][A-Za-z'’-]+){0,3}$")

    def __init__(self):
        self.cache_dir = tempfile.mkdtemp(prefix="ocpro_")

    def cleanup(self):
        shutil.rmtree(self.cache_dir, ignore_errors=True)

    def unzip_pdfs(self, zip_path: str) -> List[str]:
        out = []
        with zipfile.ZipFile(zip_path, 'r') as z:
            for m in z.infolist():
                if m.filename.lower().endswith('.pdf'):
                    dest = os.path.join(self.cache_dir, os.path.basename(m.filename))
                    z.extract(m, self.cache_dir)
                    src = os.path.join(self.cache_dir, m.filename)
                    try:
                        if os.path.abspath(src) != os.path.abspath(dest):
                            os.makedirs(os.path.dirname(dest), exist_ok=True)
                            shutil.move(src, dest)
                        out.append(dest)
                    except Exception:
                        out.append(src)
        return out

    def parse_pdf(self, pdf_path: str) -> str:
        try:
            if HAS_PDFMINER and pdfminer_extract_text:
                text = pdfminer_extract_text(pdf_path) or ""
                text = re.sub(r"\s+", " ", text)
                return text.strip()
            with open(pdf_path, "rb") as f:
                data = f.read()
            if not data.startswith(b"%PDF"):
                s = data.decode("utf-8", errors="ignore")
                return re.sub(r"\s+", " ", s).strip()
            # crude fallback
            chars, buf = [], []
            def flush():
                nonlocal buf, chars
                if buf:
                    s = bytes(buf).decode("utf-8", errors="ignore")
                    if s.strip():
                        chars.append(s)
                    buf = []
            for b in data:
                if 32 <= b <= 126 or b in (9, 10, 13) or (128 <= b <= 253):
                    buf.append(b)
                else:
                    flush()
            flush()
            text = " ".join(chars)
            text = re.sub(r"\(cid:\d+\)", " ", text)
            text = re.sub(r"\s+", " ", text)
            return text.strip()
        except Exception as e:
            return f"[PARSE_ERROR] {e}"

    def extract_contacts(self, text: str, fallback_filename: str = "") -> Tuple[str, Optional[str], Optional[str]]:
        email = self.EMAIL_RE.search(text)
        phone = self.PHONE_RE.search(text)

        # Try first ~50 tokens for a clean name
        tokens = (re.sub(r"\s+", " ", text).strip()).split()[:50]
        name: Optional[str] = None

        head = re.sub(r"[^A-Za-z '’-]", " ", " ".join(tokens))
        for ln in head.split("  "):
            ln = ln.strip()
            if not ln or "@" in ln or any(ch.isdigit() for ch in ln):
                continue
            if len(ln) <= 60 and self.NAME_LINE_RE.match(ln):
                name = ln
                break

        if not name:
            m = re.search(r"(.{0,60}?)(?:\s+Email\b|@)", text, flags=re.I)
            if m:
                guess = re.sub(r"[^A-Za-z '’-]", " ", m.group(1)).strip()
                guess = re.sub(r"\s{2,}", " ", guess)
                if 2 <= len(guess) <= 60:
                    name = guess

        if not name:
            base = os.path.splitext(os.path.basename(fallback_filename))[0]
            base = re.sub(r"[_\-]+", " ", base)
            base = re.sub(r"[^A-Za-z '’-]", " ", base).strip()
            name = base or "(Unknown)"

        return name, (email.group(0) if email else None), (phone.group(0) if phone else None)

# ---------------- Ranking -----------------
class Ranker:
    def __init__(self, keywords: Dict[str, float], role: str = ""):
        self.keywords = {k.lower(): float(v) for k, v in keywords.items()}
        self.role = role.lower()

    def score(self, text: str) -> Tuple[float, Dict[str, float]]:
        t = text.lower()
        if t.startswith("[parse_error]"):
            return 0.0, {}
        matched: Dict[str, float] = {}
        total = 0.0
        for k, w in self.keywords.items():
            cnt = t.count(k)
            if cnt:
                matched[k] = cnt * w
                total += matched[k]
        if self.role:
            total += 2.0 * t.count(self.role)
        tokens = set(re.findall(r"[a-zA-Z]{3,}", t))
        total += min(len(tokens) / 500.0, 5.0)
        return total, matched

# ---------------- Exporters -----------------
class Exporter:
    @staticmethod
    def to_csv(cands: List[Candidate], path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Name", "Email", "Phone", "Score", "Matched", "File"])
        for c in cands:
            matched = ", ".join(f"{k}:{int(v)}" for k, v in sorted(c.matched.items(), key=lambda x: -x[1]))
            with open(path, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([c.name, c.email or "", c.phone or "", round(c.score, 2), matched, os.path.basename(c.path)])

    @staticmethod
    def to_json(cands: List[Candidate], path: str):
        with open(path, "w", encoding="utf-8") as f:
            json.dump([{
                "name": c.name,
                "email": c.email,
                "phone": c.phone,
                "score": round(c.score, 2),
                "matched": c.matched,
                "file": os.path.basename(c.path),
                "summary": c.summary
            } for c in cands], f, indent=2)

    @staticmethod
    def to_pdf(cands: List[Candidate], path: str, job_title: str):
        if not HAS_REPORTLAB:
            raise RuntimeError("PDF export requires ReportLab (pip install reportlab).")
        c = canvas.Canvas(path, pagesize=A4)
        w, h = A4
        y = h - 50
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y, f"Top Candidates – {job_title}")
        y -= 30
        c.setFont("Helvetica", 10)
        for cand in cands[:50]:
            if y < 80:
                c.showPage(); y = h - 50; c.setFont("Helvetica", 10)
            c.drawString(50, y, f"{cand.name}  | Score: {round(cand.score,2)}")
            y -= 14
            c.drawString(50, y, f"Email: {cand.email or '-'}  Phone: {cand.phone or '-'}  File: {os.path.basename(cand.path)}")
            y -= 14
            c.drawString(50, y, "Matched: " + ", ".join(sorted(cand.matched.keys())) )
            y -= 20
        c.save()

# ---------------- Helpers -----------------
def open_with_default_app(path: str):
    try:
        system = platform.system()
        if system == "Darwin":
            subprocess.Popen(["open", path])
        elif system == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Open File", f"Could not open file:\n{e}")

SECTION_TITLES = [
    "Professional Summary","Summary","Objective",
    "Technical Skills","Skills","Core Skills","Tools",
    "Projects","Research & Projects","Experience","Work Experience",
    "Education","Certifications","Achievements","Awards",
    "Publications","Volunteer","Extracurricular"
]
SEC_REGEX = re.compile(r"\b(" + "|".join(re.escape(s) for s in SECTION_TITLES) + r")\b", re.I)

def prettify_blocks(raw: str, wrap_width: int = 92, char_limit: int = 7000) -> List[Tuple[str,str]]:
    """Return (tag, text) blocks for a readable preview."""
    blocks: List[Tuple[str,str]] = []
    if raw.startswith("[PARSE_ERROR]"):
        return [("para", raw)]

    txt = raw
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = SEC_REGEX.sub(r"\n\n\g<0>\n", txt)
    txt = txt.replace(" • ", "\n• ").replace(" ● ", "\n• ").replace("", "\n• ").replace("•", "\n• ")
    txt = txt.replace(" - ", " – ")
    txt = re.sub(r"\s*\n\s*", "\n", txt).strip()

    joined_len = 0
    for line in txt.split("\n"):
        s = line.strip()
        if not s:
            continue
        if SEC_REGEX.fullmatch(s):
            blocks.append(("gap",""))
            blocks.append(("sec", s))
            joined_len += len(s) + 1
            continue
        if s.startswith(("•","-","–","*")):
            plain = s.lstrip("• -*–\t").strip()
            wrapped = textwrap.fill(plain, width=wrap_width, subsequent_indent="  ")
            blocks.append(("bullet", f"• {wrapped}"))
            joined_len += len(plain) + 2
        else:
            wrapped = textwrap.fill(s, width=wrap_width)
            blocks.append(("para", wrapped))
            joined_len += len(s) + 1

        if joined_len > char_limit:
            blocks.append(("para","…"))
            break

    return blocks

# ---------------- CTk App -----------------
class CTkApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self._bg = "#000000"
        self._panel = "#0A0A0A"
        self._text = "#FFFFFF"
        self._muted = "#B3B3B3"
        self._hi = "#FFFFFF"

        # Fonts (Windows will fallback if SF Pro isn't installed)
        self._font_body = ("SF Pro Text", 12)
        self._font_body_bold = ("SF Pro Text", 12, "bold")
        self._font_heading = ("SF Pro Display", 19, "bold")
        self._font_btn = ("SF Pro Display", 13, "bold")

        # Window
        self.title(APP_NAME + " – Pro UI")
        self.geometry("1200x780")
        self.minsize(1020, 680)

        # Icon
        try:
            icon_path = resource_path("assets/app.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                self.wm_iconphoto(True, icon_img)
        except Exception:
            pass

        # If TkinterDnD base, embed CTk root
        if HAS_DND:
            self._root_ctk = ctk.CTk(master=self)
            self._root_ctk.pack(fill="both", expand=True)
            root = self._root_ctk
        else:
            root = self

        if isinstance(root, ctk.CTk):
            root.configure(fg_color=self._bg)
            root.grid_columnconfigure(1, weight=1)
            root.grid_rowconfigure(0, weight=1)

        # Licensing
        self.lic = LicenseManager()
        ok, msg = self.lic.ok()
        if not ok:
            messagebox.showwarning("License", msg)

        # Processor/data
        self.processor = CVProcessor()
        self.candidates: List[Candidate] = []
        self.filtered: List[Candidate] = []
        self.__file__queue: List[str] = []

        # Layout
        self.sidebar = ctk.CTkFrame(root, corner_radius=0, width=SIDEBAR_W, fg_color=self._panel)
        self.sidebar.grid(row=0, column=0, sticky="nswe")
        self.sidebar.grid_propagate(False)
        self.main = ctk.CTkFrame(root, corner_radius=0, fg_color=self._bg)
        self.main.grid(row=0, column=1, sticky="nswe")

        # Sidebar content
        title = ctk.CTkLabel(self.sidebar, text="CV Lense", font=("SF Pro Display", 24, "bold"), text_color=self._text)
        title.pack(padx=18, pady=(22, 4), anchor="w")

        subtitle = ctk.CTkLabel(self.sidebar, text="Offline Resume Screener", font=("SF Pro Text", 12),
                                text_color=self._muted)
        subtitle.pack(padx=18, pady=(0, 16), anchor="w")

        self.job_title = ctk.CTkEntry(
            self.sidebar,
            placeholder_text="Job title (e.g., Software Engineer)",
            text_color=self._text,
            placeholder_text_color=self._muted,
            fg_color=self._panel,
            border_color=self._hi,
            height=36
        )
        self.job_title.pack(padx=18, pady=(4, 10), fill="x")

        self.kw_var = ctk.CTkEntry(
            self.sidebar,
            placeholder_text="Keywords (kw:weight, comma separated)",
            text_color=self._text,
            placeholder_text_color=self._muted,
            fg_color=self._panel,
            border_color=self._hi,
            height=36
        )
        self.kw_var.pack(padx=18, pady=(0, 16), fill="x")

        self.drop_label = ctk.CTkLabel(
            self.sidebar,
            text=("Drop ZIP/PDFs here" if HAS_DND else "Drop ZIP/PDFs here (or Click to Browse)"),
            height=74, corner_radius=10, fg_color="#101010", text_color=self._text)
        self.drop_label.pack(padx=18, pady=(0, 10), fill="x")
        self.drop_label.bind("<Button-1>", lambda e: self.load_files_dialog())
        if HAS_DND:
            try:
                self.drop_label.drop_target_register(DND_FILES)  # type: ignore
                self.drop_label.dnd_bind("<<Drop>>", self._on_drop_paths)  # type: ignore
            except Exception:
                pass

        # Selected files list
        ctk.CTkLabel(self.sidebar, text="Selected files", text_color="#CFCFCF",
                     font=("SF Pro Text", 11, "bold")).pack(padx=18, pady=(2, 6), anchor="w")
        list_wrap = ctk.CTkFrame(self.sidebar, fg_color="#0F0F0F", corner_radius=8)
        list_wrap.pack(padx=18, pady=(0, 10), fill="both", expand=False)
        self.__file__list = tk.Listbox(list_wrap, height=6, activestyle="dotbox",
                                    bg="#0F0F0F", fg="#FFFFFF", highlightthickness=0,
                                    selectbackground="#1E1E1E", selectforeground="#FFFFFF",
                                    relief="flat", selectmode="extended")
        self.__file__list.pack(fill="both", expand=True, padx=6, pady=6)
        btn_row = ctk.CTkFrame(self.sidebar, fg_color=self._panel)
        btn_row.pack(padx=18, pady=(4, 10), fill="x")
        self.btn_add = ctk.CTkButton(btn_row, text="Add PDFs/ZIPs", command=self.load_files_dialog,
                                     fg_color="#111111", hover_color="#1A1A1A", text_color=self._text,
                                     height=34, font=("SF Pro Text", 12, "bold"), corner_radius=10)
        self.btn_add.pack(side="left", expand=True, fill="x", padx=(0, 6))
        self.btn_remove = ctk.CTkButton(btn_row, text="Remove Selected", command=self.remove_selected_from_queue,
                                        fg_color="#111111", hover_color="#1A1A1A", text_color=self._text,
                                        height=34, font=("SF Pro Text", 12, "bold"), corner_radius=10)
        self.btn_remove.pack(side="left", expand=True, fill="x")
        self.btn_clear = ctk.CTkButton(self.sidebar, text="Clear List", command=self.clear_queue,
                                       fg_color="#151515", hover_color="#1E1E1E", text_color=self._text,
                                       height=34, font=("SF Pro Text", 12, "bold"), corner_radius=10)
        self.btn_clear.pack(padx=18, pady=(0, 12), fill="x")

        self.score_btn = ctk.CTkButton(
            self.sidebar, text="Score Resumes", command=self.rank_now,
            fg_color=self._hi, text_color="#000000", hover_color="#EAEAEA",
            height=40, font=self._font_btn, corner_radius=12
        )
        self.score_btn.pack(padx=18, pady=(0, 16), fill="x")

        export_row = ctk.CTkFrame(self.sidebar, fg_color=self._panel)
        export_row.pack(padx=18, pady=(4, 10), fill="x")
        self.btn_csv = ctk.CTkButton(
            export_row, text="Export CSV", command=self.export_csv,
            fg_color="#111111", hover_color="#1A1A1A", text_color=self._text,
            height=36, font=("SF Pro Text", 12, "bold"), corner_radius=10
        )
        self.btn_csv.pack(side="left", expand=True, fill="x", padx=(0, 6))
        self.btn_json = ctk.CTkButton(
            export_row, text="Export JSON", command=self.export_json,
            fg_color="#111111", hover_color="#1A1A1A", text_color=self._text,
            height=36, font=("SF Pro Text", 12, "bold"), corner_radius=10
        )
        self.btn_json.pack(side="left", expand=True, fill="x")

        # Main
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        header = ctk.CTkLabel(self.main, text="Ranked Candidates", font=self._font_heading,
                              text_color=self._text)
        header.grid(row=0, column=0, sticky="w", padx=18, pady=(18, 10))

        table_wrap = ctk.CTkFrame(self.main, fg_color=self._bg)
        table_wrap.grid(row=1, column=0, sticky="nswe", padx=18, pady=(0, 10))
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        style = ttk.Style()
        try: style.theme_use("clam")
        except: pass

        tv_font = tkfont.Font(family="SF Pro Text", size=12)
        tv_font_bold = tkfont.Font(family="SF Pro Text", size=12, weight="bold")
        style.configure("Treeview",
                        background="#0B0B0B",
                        fieldbackground="#0B0B0B",
                        foreground=self._text,
                        rowheight=28,
                        font=tv_font)
        style.configure("Treeview.Heading",
                        background="#0F0F0F",
                        foreground=self._text,
                        font=tv_font_bold,
                        relief="flat")
        style.map("Treeview", background=[("selected", "#1E1E1E")], foreground=[("selected", "#FFFFFF")])

        cols = ("name","email","phone","score","file")
        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings", style="Treeview")
        self.tree.grid(row=0, column=0, sticky="nswe")
        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=yscroll.set)
        yscroll.grid(row=0, column=1, sticky="ns")

        self.tree.heading("name",  text="Name",  anchor="w", command=lambda: self._sort_by("name"))
        self.tree.heading("email", text="Email", anchor="w", command=lambda: self._sort_by("email"))
        self.tree.heading("phone", text="Phone", anchor="w", command=lambda: self._sort_by("phone"))
        self.tree.heading("score", text="Score", anchor="center", command=lambda: self._sort_by("score"))
        self.tree.heading("file",  text="File",  anchor="w", command=lambda: self._sort_by("file"))

        self.tree.column("name",  width=260, anchor="w", stretch=True)
        self.tree.column("email", width=220, anchor="w", stretch=False)
        self.tree.column("phone", width=150, anchor="w", stretch=False)
        self.tree.column("score", width=90,  anchor="center", stretch=False)
        self.tree.column("file",  width=240, anchor="w", stretch=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_row)
        self.tree.bind("<Double-1>", self.on_row_double_click)

        filters = ctk.CTkFrame(self.main, fg_color=self._bg)
        filters.grid(row=2, column=0, sticky="we", padx=18, pady=(0, 8))
        ctk.CTkLabel(filters, text="Search", text_color=self._text, font=self._font_body_bold).pack(side="left")
        self.search_var = tk.StringVar()
        ent_search = ctk.CTkEntry(filters, textvariable=self.search_var, width=260,
                                  text_color=self._text, fg_color="#0B0B0B", border_color=self._hi,
                                  placeholder_text="keyword, name, etc.", placeholder_text_color="#8C8C8C", height=34)
        ent_search.pack(side="left", padx=(8, 18))
        self.search_var.trace_add("write", lambda *_: self.apply_filters())

        ctk.CTkLabel(filters, text="Min Score", text_color=self._text, font=self._font_body_bold).pack(side="left")
        self.min_score_var = tk.StringVar(value="0.0")
        ent_min = ctk.CTkEntry(filters, textvariable=self.min_score_var, width=88,
                               text_color=self._text, fg_color="#0B0B0B", border_color="#0F0F0F", height=34)
        ent_min.pack(side="left", padx=(8, 18))

        self.preview = ctk.CTkTextbox(self.main, height=240, fg_color="#0B0B0B", text_color="#FFFFFF",
                                      corner_radius=8, border_width=0, wrap="word")
        self.preview.grid(row=3, column=0, sticky="we", padx=18, pady=(0, 18))

        self.license_status = tk.StringVar(value=f"License: {self.lic.ok()[1]}")
        ctk.CTkLabel(self.main, textvariable=self.license_status, text_color="#A3A3A3", font=self._font_body).grid(
            row=4, column=0, sticky="w", padx=18, pady=(0, 12)
        )

        self._pulse_phase = 0
        self._animate_pulse()
        if isinstance(root, ctk.CTk):
            root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ----- File queue helpers -----
    def update___file__list(self):
        self.__file__list.delete(0, "end")
        for p in self.__file__queue:
            self.__file__list.insert("end", os.path.basename(p))

    def remove_selected_from_queue(self):
        sel = list(self.__file__list.curselection())
        if not sel: return
        sel.reverse()
        for idx in sel:
            try:
                del self.__file__queue[idx]
            except Exception:
                pass
        self.update___file__list()
        keep = set(self.__file__queue)
        self.candidates = [c for c in self.candidates if c.path in keep]
        self.apply_filters()

    def clear_queue(self):
        self.__file__queue.clear()
        self.update___file__list()
        self.candidates.clear()
        self.apply_filters()

    # ----- Drag & Drop (multi) -----
    def _on_drop_paths(self, event):
        raw = event.data.strip()
        paths: List[str] = []
        if raw.startswith("{") and raw.endswith("}"):
            parts = re.findall(r"\{([^}]*)\}", raw)
            if parts:
                paths = parts
        else:
            paths = raw.split()
        paths = [p for p in paths if os.path.exists(p)]
        if not paths:
            messagebox.showinfo("Drop", "Please drop PDF or ZIP files.")
            return
        self._load_paths(paths)

    # ----- Add via dialog (multi) -----
    def load_files_dialog(self):
        paths = filedialog.askopenfilenames(
            title="Select PDFs and/or ZIPs",
            filetypes=[("PDF or ZIP", ".pdf *.zip"), ("PDF", ".pdf"), ("ZIP", "*.zip")]
        )
        if not paths:
            return
        self._load_paths(list(paths))

    # Core loader (APPENDS; no overwrite; processes all)
    def _load_paths(self, paths: List[str]):
        try:
            new_cache_paths: List[str] = []

            # Expand input selections to cached PDFs
            for p in paths:
                if p.lower().endswith(".zip"):
                    new_cache_paths.extend(self.processor.unzip_pdfs(p))
                elif p.lower().endswith(".pdf"):
                    dest = os.path.join(self.processor.cache_dir, os.path.basename(p))
                    try:
                        shutil.copy2(p, dest)
                        new_cache_paths.append(dest)
                    except Exception:
                        new_cache_paths.append(p)

            if not new_cache_paths:
                messagebox.showinfo("No PDFs", "No PDF files found.")
                return

            # Add to queue (avoid duplicates)
            for cp in new_cache_paths:
                if cp not in self.__file__queue:
                    self.__file__queue.append(cp)
            self.update___file__list()

            # Parse only new ones and append candidates (avoid duplicates)
            have_paths = {c.path for c in self.candidates}
            added = 0
            for path in new_cache_paths:
                if path in have_paths:
                    continue
                text = self.processor.parse_pdf(path)
                name, email, phone = self.processor.extract_contacts(text, fallback_filename=path)
                self.candidates.append(Candidate(name=name or "(Unknown)", email=email, phone=phone, path=path, text=text))
                added += 1

            total = len(self.candidates)
            messagebox.showinfo("Loaded", f"Added {added} resume(s). Total in table: {total}.")
            self.apply_filters()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load files:\n{e}")

    def parse_keywords(self) -> Dict[str,float]:
        kw = {}
        raw = self.kw_var.get()
        for token in raw.split(','):
            token = token.strip()
            if not token: continue
            if ':' in token:
                k, v = token.split(':', 1)
                try: kw[k.strip().lower()] = float(v.strip())
                except: kw[k.strip().lower()] = 1.0
            else:
                kw[token.lower()] = 1.0
        return kw

    def rank_now(self):
        if not self.candidates:
            messagebox.showinfo("No CVs", "Add PDFs or ZIPs first.")
            return
        r = Ranker(self.parse_keywords(), role=self.job_title.get())
        for c in self.candidates:
            sc, matched = r.score(c.text)
            c.score, c.matched = sc, matched
        self.candidates.sort(key=lambda x: x.score, reverse=True)
        self.apply_filters()
        if self.filtered:
            self.show_preview(self.filtered[0])

    def apply_filters(self):
        q = self.search_var.get().lower().strip()
        try:
            min_sc = float(self.min_score_var.get() or 0.0)
        except Exception:
            min_sc = 0.0
        self.filtered = [
            c for c in self.candidates
            if c.score >= min_sc and (
                q == "" or
                q in c.text.lower() or
                q in (c.name or "").lower() or
                q in os.path.basename(c.path).lower()
            )
        ]
        self.refresh_table()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for c in self.filtered:
            self.tree.insert('', 'end', values=(c.name, c.email or '', c.phone or '', round(c.score,2), os.path.basename(c.path)))

    def _sort_by(self, col: str):
        keymap = {
            "name": lambda c: (c.name or "").lower(),
            "email": lambda c: (c.email or "").lower(),
            "phone": lambda c: c.phone or "",
            "score": lambda c: c.score,
            "file": lambda c: os.path.basename(c.path).lower(),
        }
        self.filtered.sort(key=keymap[col], reverse=(col=="score"))
        self.refresh_table()

    def on_row(self, *_):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        if idx >= len(self.filtered): return
        self.show_preview(self.filtered[idx])

    def on_row_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        idx = self.tree.index(item_id)
        if idx < len(self.filtered):
            open_with_default_app(self.filtered[idx].path)

    def show_preview(self, cand: Candidate):
        self.preview.delete("1.0", "end")

        # Header lines
        self.preview.insert("end", f"{cand.name}\n", "hdr")
        self.preview.insert("end", f"Email: {cand.email or '-'}    Phone: {cand.phone or '-'}\n", "sub")
        self.preview.insert("end", f"File: {os.path.basename(cand.path)}    Score: {round(cand.score,2)}\n", "sub")
        self.preview.insert("end", "─" * 80 + "\n", "rule")

        # Matched skills summary
        topk = sorted(cand.matched.items(), key=lambda kv: -kv[1])[:10]
        if topk:
            self.preview.insert("end", "Matched Keywords\n", "sec")
            self.preview.insert("end", textwrap.fill(", ".join(k for k,_ in topk), width=92) + "\n\n")

        # Pretty content blocks
        for tag, text in prettify_blocks(cand.text):
            if tag == "gap":
                self.preview.insert("end", "\n")
            elif tag == "sec":
                self.preview.insert("end", f"{text}\n", "sec")
            elif tag == "bullet":
                self.preview.insert("end", f"{text}\n")
            else:
                self.preview.insert("end", f"{text}\n\n")

        # Styles
        try:
            self.preview.tag_configure("hdr", font=("SF Pro Display", 14, "bold"))
            self.preview.tag_configure("sub", font=("SF Pro Text", 11))
            self.preview.tag_configure("sec", font=("SF Pro Display", 12, "bold"))
            self.preview.tag_configure("rule", foreground="#666666")
        except Exception:
            pass

    def export_csv(self):
        if not self.filtered:
            return messagebox.showinfo("Nothing", "No rows to export")
        p = filedialog.asksaveasfilename(defaultextension=".csv")
        if not p: return
        Exporter.to_csv(self.filtered, p)
        messagebox.showinfo("Saved", p)

    def export_json(self):
        if not self.filtered:
            return messagebox.showinfo("Nothing", "No rows to export")
        p = filedialog.asksaveasfilename(defaultextension=".json")
        if not p: return
        Exporter.to_json(self.filtered, p)
        messagebox.showinfo("Saved", p)

    def _on_close(self):
        try:
            self.processor.cleanup()
        finally:
            try:
                self.destroy()
            except Exception:
                pass

    # Pulse animation
    def _animate_pulse(self):
        self._pulse_phase = (getattr(self, "_pulse_phase", 0) + 1) % 60
        phase = self._pulse_phase / 60.0
        lum = 0.5 + 0.5 * math.cos(2 * math.pi * phase)
        a = (255, 255, 255)
        b = (234, 234, 234)
        rgb = tuple(int(a[i] * (1 - lum * 0.35) + b[i] * (lum * 0.35)) for i in range(3))
        color = "#{:02X}{:02X}{:02X}".format(*rgb)
        try:
            self.score_btn.configure(fg_color=color, hover_color="#EAEAEA", text_color="#000000")
        except Exception:
            pass
        self.after(90, self._animate_pulse)

# ---------------- Main -----------------
if __name__ == "__main__":
    app = CTkApp()
    app.mainloop()
