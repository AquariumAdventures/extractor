# -------------------------------------------------------------
# Extractor – ChatGPT Vision Backend (Tkinter GUI)
# -------------------------------------------------------------
# • Requires: `openai>=1.10.0`  ➜  `pip install openai requests pillow`
# • Environment: set `OPENAI_API_KEY`
# • Optional clipboard: `pip install pyperclip`
# • Icon: put `icon.ico` next to the script / EXE
# -------------------------------------------------------------

import base64
import csv
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from io import BytesIO

import requests
from PIL import Image, ImageTk
import openai  # core dependency

try:
    import pyperclip  # optional – clipboard copy
except ImportError:
    pyperclip = None

# --------------------------- OpenAI settings ---------------------------
MODEL = "gpt-4o-mini"  # change to gpt-4o if your account has access
PROMPT_IMAGE = (
    "You are given an image that contains tabular data. "
    "Identify sensible column names and extract the table. "
    "Return ONLY raw CSV with a header row and subsequent data rows."
)
# ----------------------------------------------------------------------

# keep reference to the thumbnail so it isn't garbage‑collected
_thumbnail_ref = None

# --------------------------- OpenAI helpers ---------------------------

def send_to_openai(image_bytes: bytes) -> str:
    """Send image bytes to GPT‑4o Vision and return raw CSV."""
    b64 = base64.b64encode(image_bytes).decode()
    resp = openai.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT_IMAGE},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            }
        ],
    )
    return resp.choices[0].message.content.strip()

# --------------------------- CSV helpers ---------------------------

def parse_csv(csv_text: str):
    lines = []
    for line in csv_text.splitlines():
        line = line.strip()
        if not line or line.startswith("```"):
            continue
        lines.append(line)
    return [row for row in csv.reader(lines) if any(cell.strip() for cell in row)]


def save_and_copy(rows, out_dir):
    out_path = os.path.join(out_dir, "results.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    if pyperclip:
        try:
            pyperclip.copy("\n".join([",".join(r) for r in rows]))
        except Exception:
            pass
    try:
        os.startfile(out_path)
    except Exception:
        pass
    return out_path

# --------------------------- UI helpers ---------------------------

def choose_headers(root, headers, preview_rows):
    result = []
    dlg = tk.Toplevel(root)
    dlg.title("Select Columns to Export")

    ttk.Label(dlg, text="Tick columns to include:").pack(anchor="w", padx=10, pady=(10, 0))

    preview = ttk.Treeview(dlg, columns=headers, show="headings", height=min(5, len(preview_rows)))
    for c in headers:
        preview.heading(c, text=c)
        preview.column(c, width=120)
    for r in preview_rows:
        preview.insert("", "end", values=r)
    preview.pack(fill="both", expand=True, padx=10, pady=5)

    vars_ = []
    for h in headers:
        var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dlg, text=h, variable=var).pack(anchor="w", padx=10)
        vars_.append((h, var))

    btn = ttk.Button(dlg, text="OK", command=dlg.destroy)
    btn.pack(pady=10)

    dlg.grab_set()
    dlg.wait_window()

    for h, v in vars_:
        if v.get():
            result.append(h)
    return result

# --------------------------- main logic ---------------------------

def run_extraction(app, img_path):
    """Background worker thread."""
    try:
        with open(img_path, "rb") as f:
            bytes_ = f.read()
        csv_raw = send_to_openai(bytes_)
        rows = parse_csv(csv_raw)
        if not rows:
            raise ValueError("No rows parsed – model response seems empty or malformed.")
        headers = rows[0]
        preview_rows = rows[1:6]
        selected = choose_headers(app.root, headers, preview_rows)
        if not selected:
            raise ValueError("No columns selected.")
        idx = [i for i, h in enumerate(headers) if h in selected]
        filtered = [[row[i] if i < len(row) else "" for i in idx] for row in rows]
        save_path = save_and_copy(filtered, os.path.dirname(img_path))
        # update table
        app.table.config(columns=selected)
        for c in selected:
            app.table.heading(c, text=c)
            app.table.column(c, width=120)
        for r in filtered[1:]:
            app.table.insert("", "end", values=r)
        app.set_status(f"Done – saved to {save_path}")
    except Exception as e:
        app.set_status("Error – see dialog")
        messagebox.showerror("Extraction failed", str(e))
    finally:
        app.progress.stop()
        app.progress.place_forget()

# --------------------------- App class ---------------------------

class ExtractorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Extractor")
        self.root.geometry("640x520")
        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        # dark mode toggle placeholder – could add switch later

    def _build_ui(self):
        frm_top = ttk.Frame(self.root, padding=10)
        frm_top.pack(fill="x")
        ttk.Label(frm_top, text="Image file with a table:").pack(anchor="w")

        self.path_var = tk.StringVar()
        ent = ttk.Entry(frm_top, textvariable=self.path_var, width=60)
        ent.pack(side="left", fill="x", expand=True)
        ttk.Button(frm_top, text="Browse", command=self._browse).pack(side="right")

        # thumbnail preview
        self.thumb_lbl = ttk.Label(self.root)
        self.thumb_lbl.pack(pady=5)

        # action buttons
        frm_act = ttk.Frame(self.root, padding=(10, 5))
        frm_act.pack()
        ttk.Button(frm_act, text="Extract", command=self._on_extract).pack(side="left", padx=5)
        ttk.Button(frm_act, text="Exit", command=self.root.quit).pack(side="left")

        # table view
        self.table = ttk.Treeview(self.root, show="headings")
        self.table.pack(fill="both", expand=True, padx=10, pady=10)

        # indeterminate progress bar
        self.progress = ttk.Progressbar(self.root, mode="indeterminate")

        # status bar
        self.status_var = tk.StringVar()
        ttk.Label(self.root, textvariable=self.status_var, anchor="w").pack(fill="x", side="bottom")

    # -------------- helpers --------------
    def set_status(self, msg):
        self.status_var.set(msg)

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Image", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
        if p:
            self.path_var.set(p)
            self._show_thumbnail(p)

    def _show_thumbnail(self, path):
        global _thumbnail_ref
        try:
            img = Image.open(path)
            img.thumbnail((200, 200))
            _thumbnail_ref = ImageTk.PhotoImage(img)
            self.thumb_lbl.configure(image=_thumbnail_ref)
        except Exception:
            self.thumb_lbl.configure(image="")

    def _on_extract(self):
        path = self.path_var.get().strip()
        if not os.path.isfile(path):
            messagebox.showerror("No image", "Please choose an image file first.")
            return
        self.table.delete(*self.table.get_children())
        self.set_status("Starting extraction …")
        # show progress bar
        self.progress.pack(fill="x", padx=10)
        self.progress.start(12)
        threading.Thread(target=run_extraction, args=(self, path), daemon=True).start()

# --------------------------- run ---------------------------
if __name__ == "__main__":
    # Prompt for API key if not set in the environment (key is only kept for this session)
    if not os.getenv("OPENAI_API_KEY"):
        from tkinter import simpledialog
        tk_popup = tk.Tk()
        tk_popup.withdraw()
        user_key = simpledialog.askstring(
            "Enter API Key", "Please enter your OpenAI API key:", show="*")
        tk_popup.destroy()
        if not user_key:
            messagebox.showerror("No Key Entered", "The application cannot run without an API key.")
            sys.exit(1)
        os.environ["OPENAI_API_KEY"] = user_key

    # Launch the main application UI
    app = ExtractorApp()
    app.root.mainloop()
