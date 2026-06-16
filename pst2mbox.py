#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PST → Thunderbird Converter  (v3 — pypff con sub_items fallback)
Convierte archivos .pst de Outlook (2007-2026) a formato Mbox para Thunderbird.

Instalación:
  Linux/WSL : sudo apt install python3-pypff
  Windows   : pip install libpff-python
  macOS     : brew install libpff && pip install libpff-python

Empaquetado:
  pyinstaller --onefile --noconsole pst2mbox.py
  pyinstaller --noconsole --onefile --add-binary "RUTA_A_PYPFF;pypff" pst2mbox.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import sys
import email.utils
import base64
import quopri
import time
import re
from pathlib import Path
from datetime import datetime

# ─── Constantes MAPI ─────────────────────────────────────────────────────────
MAPI_ATTACH_LONG_FILENAME = 0x3707
MAPI_ATTACH_FILENAME      = 0x3704
MAPI_ATTACH_MIME_TAG      = 0x370E
MAPI_SENDER_EMAIL         = 0x0C1F
MAPI_SENDER_EMAIL_ALT     = 0x0065
MAPI_TO_RECIPIENTS        = 0x0E04
MAPI_CC_RECIPIENTS        = 0x0E03
MAPI_MESSAGE_ID           = 0x1035
MAPI_IN_REPLY_TO          = 0x1042

# ─── Paleta visual ────────────────────────────────────────────────────────────
BG        = "#0f1117"
BG2       = "#181c27"
BG3       = "#1e2335"
CARD      = "#232840"
ACCENT    = "#5b8af5"
ACCENT2   = "#7c63f0"
SUCCESS   = "#3ecf8e"
WARNING   = "#f5a623"
ERROR_COL = "#f55b5b"
TEXT      = "#e8ecf5"
TEXT2     = "#8892a4"
BORDER    = "#2d3452"

FONT_TITLE  = ("Segoe UI", 18, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)


# ─── Comprobación pypff ───────────────────────────────────────────────────────

def check_pypff():
    try:
        import pypff
        print(pypff.__file__)
        return True, f"pypff {pypff.get_version()}"
    except ImportError:
        return False, ""

def install_instructions():
    if sys.platform.startswith("linux"):
        return "  sudo apt install python3-pypff\n  (o: pip install libpff-python)"
    elif sys.platform == "darwin":
        return "  brew install libpff\n  pip install libpff-python"
    else:
        return "  pip install libpff-python\n  (requiere Visual Studio Build Tools)"


# ─── Núcleo de conversión ─────────────────────────────────────────────────────

class ConversionStats:
    def __init__(self):
        self.folders = self.emails = self.attachments = self.errors = 0
        self.size_in = self.size_out = 0
        self.start_time = self.end_time = None

    def elapsed(self):
        if not self.start_time:
            return "0s"
        t = (self.end_time or time.time()) - self.start_time
        return f"{t:.1f}s" if t < 60 else f"{int(t//60)}m {int(t%60)}s"


def _safe_str(value, fallback=""):
    if value is None:
        return fallback
    if isinstance(value, bytes):
        for enc in ("utf-8", "latin-1", "cp1252"):
            try:
                return value.decode(enc)
            except Exception:
                continue
        return value.decode("latin-1", errors="replace")
    return str(value)


def _get_mapi_prop(item, prop_id):
    try:
        for rs_idx in range(item.number_of_record_sets):
            rs = item.get_record_set(rs_idx)
            for e_idx in range(rs.number_of_entries):
                entry = rs.get_entry(e_idx)
                if entry.entry_type == prop_id:
                    try:
                        return entry.data_as_string
                    except Exception:
                        try:
                            raw = entry.data
                            return _safe_str(raw) if raw else None
                        except Exception:
                            return None
    except Exception:
        pass
    return None


def _format_date(dt):
    try:
        return email.utils.formatdate(dt.timestamp(), localtime=False)
    except Exception:
        return email.utils.formatdate(localtime=False)


def _encode_qp(text):
    if not text:
        return ""
    try:
        return quopri.encodestring(text.encode("utf-8")).decode("ascii")
    except Exception:
        return text


def _safe_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name)
    return name[:200] or "attachment"


def _safe_foldername(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', _safe_str(name, "carpeta"))
    name = name.strip(". ")
    return name[:100] or "carpeta"


def _get_attach_name(att):
    for pid in (MAPI_ATTACH_LONG_FILENAME, MAPI_ATTACH_FILENAME):
        n = _get_mapi_prop(att, pid)
        if n and n.strip():
            return n.strip()
    return f"adjunto_{att.identifier}"


def _get_attach_mime(att):
    m = _get_mapi_prop(att, MAPI_ATTACH_MIME_TAG)
    return m.strip() if m and "/" in m else "application/octet-stream"


def _item_is_message(item):
    """Detecta si un sub_item es realmente un mensaje de correo."""
    import pypff
    if isinstance(item, pypff.message):
        return True
    # Fallback: intentar leer subject (los mensajes lo tienen)
    try:
        _ = item.subject
        return True
    except Exception:
        pass
    try:
        _ = item.sender_name
        return True
    except Exception:
        pass
    return False


def _iter_messages(folder):
    """
    Genera todos los mensajes de una carpeta probando primero sub_messages
    y si no hay, recurriendo a sub_items filtrando los que son mensajes.
    """
    import pypff

    # Vía 1: sub_messages (API directa)
    try:
        n = folder.number_of_sub_messages
        if n and n > 0:
            for i in range(n):
                try:
                    yield folder.get_sub_message(i)
                except Exception:
                    pass
            return  # si funcionó, no necesitamos sub_items
    except Exception:
        pass

    # Vía 2: sub_items (muchos PST usan esto)
    try:
        n = folder.number_of_sub_items
        if n and n > 0:
            for i in range(n):
                try:
                    item = folder.get_sub_item(i)
                    if _item_is_message(item):
                        yield item
                except Exception:
                    pass
    except Exception:
        pass


def _count_messages(folder):
    """Cuenta mensajes usando ambas vías."""
    try:
        n = folder.number_of_sub_messages
        if n and n > 0:
            return n
    except Exception:
        pass
    try:
        n = folder.number_of_sub_items
        if n and n > 0:
            count = 0
            for i in range(n):
                try:
                    item = folder.get_sub_item(i)
                    if _item_is_message(item):
                        count += 1
                except Exception:
                    pass
            return count
    except Exception:
        pass
    return 0


def _build_mbox_message(msg, opts):
    """Construye un mensaje mbox desde un item pypff. Devuelve (bytes, n_adjuntos)."""
    subject     = _safe_str(msg.subject, "(Sin asunto)")
    sender_name = _safe_str(msg.sender_name, "")
    sender_mail = (_get_mapi_prop(msg, MAPI_SENDER_EMAIL) or
                   _get_mapi_prop(msg, MAPI_SENDER_EMAIL_ALT) or
                   "unknown@unknown")

    dt = None
    try:
        dt = msg.delivery_time or msg.client_submit_time
    except Exception:
        pass

    date_str  = _format_date(dt)
    from_line = f"From {sender_mail} {email.utils.formatdate(localtime=False)}\n"

    html_body = plain_body = None
    if opts.get("prefer_html"):
        try:
            h = msg.html_body
            if h and len(h) > 5:
                html_body = h
        except Exception:
            pass
    try:
        p = msg.plain_text_body
        if p and len(p) > 1:
            plain_body = p
    except Exception:
        pass

    to_str   = _get_mapi_prop(msg, MAPI_TO_RECIPIENTS) or ""
    cc_str   = _get_mapi_prop(msg, MAPI_CC_RECIPIENTS) or ""
    msg_id   = _get_mapi_prop(msg, MAPI_MESSAGE_ID)    or f"<pst.{msg.identifier}@converted>"
    reply_to = _get_mapi_prop(msg, MAPI_IN_REPLY_TO)   or ""

    # Adjuntos
    attachments_data = []
    if opts.get("include_attachments"):
        try:
            n_att = msg.number_of_attachments
            for i in range(n_att or 0):
                try:
                    att      = msg.get_attachment(i)
                    att_size = att.size or 0
                    if att_size > 0:
                        att.seek_offset(0)
                        data = att.read_buffer(att_size)
                        if data:
                            attachments_data.append(
                                (_get_attach_name(att), _get_attach_mime(att), data))
                except Exception:
                    pass
        except Exception:
            pass

    boundary = f"----=_Part_{msg.identifier}_{int(time.time()*1000)}"
    has_att  = bool(attachments_data)
    has_both = html_body and plain_body

    lines = [from_line,
             f"From: {sender_name} <{sender_mail}>\n",
             f"To: {to_str}\n"]
    if cc_str:
        lines.append(f"Cc: {cc_str}\n")
    lines += [f"Subject: {subject}\n",
              f"Date: {date_str}\n",
              f"Message-ID: {msg_id}\n"]
    if reply_to:
        lines.append(f"In-Reply-To: {reply_to}\n")
    lines.append("MIME-Version: 1.0\n")
    lines.append("X-Converted-By: PST-Thunderbird-Converter-v3\n")

    if has_att or has_both:
        lines.append(f"Content-Type: multipart/mixed; boundary=\"{boundary}\"\n\n")
        inner = f"----=_Inner_{msg.identifier}"

        if has_both:
            lines += [f"--{boundary}\n",
                      f"Content-Type: multipart/alternative; boundary=\"{inner}\"\n\n",
                      f"--{inner}\n",
                      "Content-Type: text/plain; charset=utf-8\n",
                      "Content-Transfer-Encoding: quoted-printable\n\n",
                      _encode_qp(_safe_str(plain_body)), "\n",
                      f"--{inner}\n",
                      "Content-Type: text/html; charset=utf-8\n",
                      "Content-Transfer-Encoding: quoted-printable\n\n",
                      _encode_qp(_safe_str(html_body)), "\n",
                      f"--{inner}--\n"]
        elif html_body:
            lines += [f"--{boundary}\n",
                      "Content-Type: text/html; charset=utf-8\n",
                      "Content-Transfer-Encoding: quoted-printable\n\n",
                      _encode_qp(_safe_str(html_body)), "\n"]
        else:
            lines += [f"--{boundary}\n",
                      "Content-Type: text/plain; charset=utf-8\n",
                      "Content-Transfer-Encoding: quoted-printable\n\n",
                      _encode_qp(_safe_str(plain_body or "")), "\n"]

        for att_name, att_mime, att_data in attachments_data:
            safe  = _safe_filename(att_name)
            lines += [f"--{boundary}\n",
                      f"Content-Type: {att_mime}; name=\"{safe}\"\n",
                      "Content-Transfer-Encoding: base64\n",
                      f"Content-Disposition: attachment; filename=\"{safe}\"\n\n",
                      base64.encodebytes(att_data).decode("ascii"), "\n"]
        lines.append(f"--{boundary}--\n\n")

    else:
        body_text    = _safe_str(html_body or plain_body or "")
        content_type = "text/html" if html_body else "text/plain"
        lines += [f"Content-Type: {content_type}; charset=utf-8\n",
                  "Content-Transfer-Encoding: quoted-printable\n\n",
                  _encode_qp(body_text), "\n\n"]

    return "".join(lines).encode("utf-8", errors="replace"), len(attachments_data)


def process_folder(folder, out_path, opts, stats, log_cb, depth=0):
    folder_name = _safe_str(folder.name, "Sin_nombre")
    safe_name   = _safe_foldername(folder_name)
    indent      = "  " * depth

    n_msgs = _count_messages(folder)
    log_cb(f"{indent}📁 {folder_name} ", "plain")

    if n_msgs > 0:
        mbox_path = out_path / f"{safe_name}.mbox"
        written = att_count = 0
        try:
            with open(mbox_path, "wb") as mbox_file:
                for msg in _iter_messages(folder):
                    try:
                        msg_bytes, n_att = _build_mbox_message(msg, opts)
                        mbox_file.write(msg_bytes)
                        written   += 1
                        att_count += n_att
                        stats.size_out += len(msg_bytes)
                    except Exception as e:
                        stats.errors += 1
                        log_cb(f"\n{indent}  ⚠ mensaje: {e}\n", "warn")
            stats.emails      += written
            stats.attachments += att_count
            log_cb(f"({written} mensajes, {att_count} adjuntos)\n", "ok")
        except Exception as e:
            stats.errors += 1
            log_cb(f"\n{indent}  ✗ Error mbox: {e}\n", "error")
    else:
        log_cb("(vacía)\n", "info")

    # Subcarpetas
    try:
        n_sub = folder.number_of_sub_folders
        if n_sub and n_sub > 0:
            sub_out = out_path / safe_name
            sub_out.mkdir(parents=True, exist_ok=True)
            stats.folders += 1
            for i in range(n_sub):
                try:
                    process_folder(folder.get_sub_folder(i),
                                   sub_out, opts, stats, log_cb, depth + 1)
                except Exception as e:
                    stats.errors += 1
                    log_cb(f"{indent}  ✗ Subcarpeta {i}: {e}\n", "error")
    except Exception:
        pass


def run_conversion(pst_file, out_dir, opts, log_cb, progress_cb, done_cb):
    import pypff

    stats            = ConversionStats()
    stats.start_time = time.time()
    pst_path         = Path(pst_file)
    out_path         = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    stats.size_in    = pst_path.stat().st_size

    log_cb(f"  Archivo  : {pst_path.name}\n", "info")
    log_cb(f"  Tamaño   : {stats.size_in/1024/1024:.1f} MB\n", "info")
    log_cb(f"  Destino  : {out_path}\n", "info")
    log_cb(f"  Motor    : pypff {pypff.get_version()}\n", "info")
    log_cb("─"*48 + "\n\n", "sep")

    try:
        pff = pypff.file()
        pff.open(str(pst_path))
        log_cb("  ✓ PST abierto correctamente.\n\n", "ok")
        progress_cb(5)

        root  = pff.get_root_folder()
        n_top = root.number_of_sub_folders
        log_cb(f"  Carpetas raíz: {n_top}\n\n", "info")

        for i in range(n_top):
            try:
                process_folder(root.get_sub_folder(i),
                               out_path, opts, stats, log_cb, depth=1)
                stats.folders += 1
                progress_cb(min(10 + int((i+1)/n_top*85), 95))
            except Exception as e:
                stats.errors += 1
                log_cb(f"  ✗ Carpeta raíz {i}: {e}\n", "error")

        pff.close()
        stats.end_time = time.time()
        progress_cb(100)

        log_cb("─"*48 + "\n", "sep")
        log_cb("  RESUMEN\n", "header")
        log_cb("─"*48 + "\n", "sep")
        log_cb(f"  • Carpetas   : {stats.folders}\n", "stat")
        log_cb(f"  • Correos    : {stats.emails}\n", "stat")
        log_cb(f"  • Adjuntos   : {stats.attachments}\n", "stat")
        log_cb(f"  • Errores    : {stats.errors}\n",
               "stat_warn" if stats.errors else "stat")
        log_cb(f"  • Entrada    : {stats.size_in/1024/1024:.1f} MB\n", "stat")
        log_cb(f"  • Salida     : {stats.size_out/1024/1024:.1f} MB\n", "stat")
        log_cb(f"  • Tiempo     : {stats.elapsed()}\n", "stat")
        log_cb("─"*48 + "\n\n", "sep")
        log_cb(f"  {out_path}\n", "path")

        done_cb(True, stats)

    except Exception as e:
        stats.end_time = time.time()
        log_cb(f"\n  ✗ Error fatal: {e}\n", "error")
        import traceback
        log_cb(traceback.format_exc(), "error")
        done_cb(False, stats)


# ─── GUI ─────────────────────────────────────────────────────────────────────

class PSTConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PST → Thunderbird Converter  v3")
        self.geometry("900x780")
        self.minsize(780, 660)
        self.configure(bg=BG)
        self._center()
        self._styles()
        self._build_ui()
        self._check_on_start()

    def _center(self):
        self.update_idletasks()
        w, h = 900, 780
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",       background=BG)
        s.configure("TLabel",       background=BG, foreground=TEXT, font=FONT_BODY)
        s.configure("TCheckbutton", background=CARD, foreground=TEXT,
                    font=FONT_BODY, selectcolor=BG3)
        s.map("TCheckbutton",
              background=[("active", CARD)], foreground=[("active", TEXT)])
        s.configure("Accent.TButton",
                    background=ACCENT, foreground="white",
                    font=("Segoe UI", 10, "bold"),
                    borderwidth=0, relief="flat", padding=(16, 8))
        s.map("Accent.TButton",
              background=[("active", ACCENT2), ("disabled", BORDER)])
        s.configure("Secondary.TButton",
                    background=BG3, foreground=TEXT2,
                    font=FONT_BODY, borderwidth=0, relief="flat", padding=(10, 6))
        s.map("Secondary.TButton",
              background=[("active", CARD)], foreground=[("active", TEXT)])
        s.configure("TProgressbar",
                    background=ACCENT, troughcolor=BG3, borderwidth=0, thickness=6)

    def _build_ui(self):
        banner = tk.Frame(self, bg=BG2, height=70)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="✉  PST → Thunderbird",
                 bg=BG2, fg=TEXT, font=FONT_TITLE).pack(side="left", padx=24, pady=10)
        self.status_badge = tk.Label(banner, text="⬤  Comprobando…",
                                     bg=BG2, fg=WARNING, font=FONT_SMALL)
        self.status_badge.pack(side="right", padx=20)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)
        left = tk.Frame(body, bg=BG, width=430)
        left.pack(side="left", fill="both", expand=True, padx=(16,8), pady=12)
        left.pack_propagate(False)
        right = tk.Frame(body, bg=BG)
        right.pack(side="right", fill="both", expand=True, padx=(8,16), pady=12)

        self._build_left(left)
        self._build_right(right)

        bar = tk.Frame(self, bg=BG2, height=32)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.footer_lbl = tk.Label(bar, text="Listo.", bg=BG2, fg=TEXT2, font=FONT_SMALL)
        self.footer_lbl.pack(side="left", padx=16, pady=6)

    def _card(self, parent, title):
        w = tk.Frame(parent, bg=CARD)
        w.pack(fill="x", pady=(0, 10))
        tk.Label(w, text=title, bg=CARD, fg=TEXT2,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14, pady=(10,2))
        inner = tk.Frame(w, bg=CARD)
        inner.pack(fill="x", padx=14, pady=(0,12))
        return inner

    def _entry_row(self, parent, var, cmd):
        r = tk.Frame(parent, bg=CARD)
        r.pack(fill="x")
        tk.Entry(r, textvariable=var, bg=BG3, fg=TEXT, insertbackground=TEXT,
                 font=FONT_BODY, relief="flat",
                 highlightthickness=1, highlightcolor=ACCENT,
                 highlightbackground=BORDER
                 ).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        ttk.Button(r, text="Examinar…", style="Secondary.TButton",
                   command=cmd).pack(side="right")

    def _build_left(self, parent):
        c1 = self._card(parent, "ARCHIVO PST DE ENTRADA")
        self.pst_var = tk.StringVar()
        self._entry_row(c1, self.pst_var, self._browse_pst)
        self.pst_info = tk.Label(c1, text="", bg=CARD, fg=TEXT2, font=FONT_SMALL)
        self.pst_info.pack(anchor="w", pady=(4,0))

        c2 = self._card(parent, "CARPETA DE DESTINO")
        self.out_var = tk.StringVar()
        self._entry_row(c2, self.out_var, self._browse_out)
        tk.Label(c2, text="Se creará si no existe. Los mbox previos serán sobreescritos.",
                 bg=CARD, fg=TEXT2, font=FONT_SMALL).pack(anchor="w", pady=(4,0))

        c3 = self._card(parent, "OPCIONES")
        self.opt_att  = tk.BooleanVar(value=True)
        self.opt_html = tk.BooleanVar(value=True)
        for var, label, desc in [
            (self.opt_att,  "Incluir adjuntos e imágenes", "Embebe en base64 en los mbox"),
            (self.opt_html, "Preferir cuerpo HTML",        "Usa HTML si está disponible"),
        ]:
            row = tk.Frame(c3, bg=CARD); row.pack(fill="x", pady=2)
            tk.Checkbutton(row, variable=var, text=label,
                           bg=CARD, fg=TEXT, selectcolor=BG3,
                           activebackground=CARD, activeforeground=TEXT,
                           font=FONT_BODY, cursor="hand2",
                           highlightthickness=0).pack(side="left")
            tk.Label(row, text=f"  {desc}",
                     bg=CARD, fg=TEXT2, font=FONT_SMALL).pack(side="left")

        bf = tk.Frame(parent, bg=BG); bf.pack(fill="x", pady=(4,0))
        self.convert_btn = ttk.Button(bf, text="▶  Convertir",
                                      style="Accent.TButton", command=self._start)
        self.convert_btn.pack(side="left", fill="x", expand=True, padx=(0,6))
        ttk.Button(bf, text="↺  Limpiar", style="Secondary.TButton",
                   command=self._reset).pack(side="right")

        pf = tk.Frame(parent, bg=BG); pf.pack(fill="x", pady=(10,0))
        self.progress_var = tk.IntVar(value=0)
        ttk.Progressbar(pf, variable=self.progress_var,
                        maximum=100, style="TProgressbar").pack(fill="x")
        pr = tk.Frame(pf, bg=BG); pr.pack(fill="x")
        self.pct_lbl = tk.Label(pr, text="", bg=BG, fg=TEXT2, font=FONT_SMALL)
        self.pct_lbl.pack(side="right")

    def _build_right(self, parent):
        hr = tk.Frame(parent, bg=BG); hr.pack(fill="x", pady=(0,6))
        tk.Label(hr, text="Registro de actividad",
                 bg=BG, fg=TEXT2,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        ttk.Button(hr, text="Limpiar log", style="Secondary.TButton",
                   command=self._clear_log).pack(side="right")

        lf = tk.Frame(parent, bg=CARD); lf.pack(fill="both", expand=True)
        self.log = tk.Text(lf, bg=BG2, fg=TEXT, font=FONT_MONO, relief="flat",
                           wrap="none", state="disabled", padx=10, pady=8,
                           selectbackground=ACCENT, insertbackground=TEXT,
                           highlightthickness=0, spacing1=1, spacing3=1)
        vsb = tk.Scrollbar(lf, orient="vertical",   command=self.log.yview, width=8)
        hsb = tk.Scrollbar(lf, orient="horizontal", command=self.log.xview, width=8)
        self.log.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.log.pack(fill="both", expand=True)

        for tag, cfg in {
            "info":      {"foreground": TEXT2},
            "ok":        {"foreground": SUCCESS},
            "warn":      {"foreground": WARNING},
            "error":     {"foreground": ERROR_COL},
            "plain":     {"foreground": TEXT},
            "header":    {"foreground": ACCENT, "font": ("Consolas", 9, "bold")},
            "stat":      {"foreground": TEXT},
            "stat_warn": {"foreground": WARNING},
            "sep":       {"foreground": BORDER},
            "path":      {"foreground": ACCENT2},
        }.items():
            self.log.tag_configure(tag, **cfg)

        self._log("  PST → Thunderbird Converter  v3\n", "header")
        self._log(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", "info")
        self._log("─"*46 + "\n\n", "sep")

    def _check_on_start(self):
        ok, ver = check_pypff()
        if ok:
            self.status_badge.configure(text=f"⬤  {ver}", fg=SUCCESS)
            self._log(f"  ✓ {ver} detectado.\n\n", "ok")
            self._log("  1. Selecciona el archivo .pst\n", "plain")
            self._log("  2. Elige carpeta de destino\n", "plain")
            self._log("  3. Pulsa ▶ Convertir\n\n", "plain")
        else:
            self.status_badge.configure(text="⬤  pypff NO encontrado", fg=ERROR_COL)
            self._log("[✗] pypff no instalado.\n\n", "error")
            self._log(install_instructions() + "\n", "warn")
            self.convert_btn.configure(state="disabled")

    def _browse_pst(self):
        p = filedialog.askopenfilename(
            title="Seleccionar archivo PST",
            filetypes=[("Archivos PST", "*.pst"), ("Todos", "*.*")])
        if p:
            self.pst_var.set(p)
            size = Path(p).stat().st_size / 1024 / 1024
            self.pst_info.configure(text=f"Tamaño: {size:.1f} MB  |  {Path(p).name}")
            if not self.out_var.get():
                self.out_var.set(str(Path(p).parent / (Path(p).stem + "_thunderbird")))
            self._log(f"  PST: {Path(p).name}  ({size:.1f} MB)\n", "info")

    def _browse_out(self):
        p = filedialog.askdirectory(title="Carpeta de destino")
        if p:
            self.out_var.set(p)
            self._log(f"  Destino: {p}\n", "info")

    def _start(self):
        pst = self.pst_var.get().strip()
        out = self.out_var.get().strip()
        if not pst:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PST."); return
        if not Path(pst).exists():
            messagebox.showerror("No encontrado", f"No existe:\n{pst}"); return
        if not out:
            messagebox.showwarning("Falta destino", "Indica la carpeta de destino."); return

        out_p = Path(out)
        if out_p.exists() and any(out_p.iterdir()):
            if not messagebox.askyesno("Carpeta existente",
                    f"'{out}' ya tiene archivos.\n¿Continuar y sobreescribir?"):
                return

        opts = {"include_attachments": self.opt_att.get(),
                "prefer_html":         self.opt_html.get()}

        self._clear_log()
        self.convert_btn.configure(state="disabled", text="⏳  Convirtiendo…")
        self.progress_var.set(0)
        self.pct_lbl.configure(text="")
        self.footer_lbl.configure(text="Convirtiendo… por favor espera.", fg=TEXT2)
        self._log("─"*46 + "\n", "sep")
        self._log(f"  INICIO: {datetime.now().strftime('%H:%M:%S')}\n", "header")
        self._log("─"*46 + "\n\n", "sep")

        threading.Thread(
            target=run_conversion,
            args=(pst, out, opts, self._log_ts, self._prog_ts, self._done_ts),
            daemon=True
        ).start()

    def _reset(self):
        self.pst_var.set(""); self.out_var.set("")
        self.pst_info.configure(text="")
        self.progress_var.set(0); self.pct_lbl.configure(text="")
        self.footer_lbl.configure(text="Listo.", fg=TEXT2)
        self._clear_log()

    def _log(self, text, tag="plain"):
        self.log.configure(state="normal")
        self.log.insert("end", text, tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _log_ts(self, text, tag="plain"):
        self.after(0, self._log, text, tag)

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _prog_ts(self, val):
        self.after(0, lambda: (self.progress_var.set(val),
                               self.pct_lbl.configure(text=f"{val}%")))

    def _done_ts(self, success, stats):
        def _u():
            self.convert_btn.configure(state="normal", text="▶  Convertir")
            if success:
                self.footer_lbl.configure(
                    text=f"✓ {stats.elapsed()}  |  {stats.emails} correos  |  {stats.attachments} adjuntos",
                    fg=SUCCESS)
                messagebox.showinfo("Completado",
                    f"✅ Conversión finalizada.\n\n"
                    f"📁 Carpetas : {stats.folders}\n"
                    f"✉️  Correos  : {stats.emails}\n"
                    f"📎 Adjuntos : {stats.attachments}\n"
                    f"⏱️  Tiempo   : {stats.elapsed()}\n\n"
                    f"Destino:\n{self.out_var.get()}\n\n"
                    "Para importar en Thunderbird:\n"
                    "  Addon: ImportExportTools NG\n"
                    "  → Importar carpeta mbox")
            else:
                self.footer_lbl.configure(
                    text="✗ Error en la conversión. Revisa el log.", fg=ERROR_COL)
        self.after(0, _u)

# Antes de ejecutar instalar pypff con pip install libpff-python (requiere Visual Studio Build Tools en Windows)
if __name__ == "__main__":
    app = PSTConverter()
    app.mainloop()