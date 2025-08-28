import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os, json, smtplib, mimetypes, re, threading, socket, html
from email.message import EmailMessage

# =============== Política de credenciales ===============
ALLOW_PERSIST = False  # no persistir credenciales

# =============== Opcional: keyring ===============
try:
    import keyring
except Exception:
    keyring = None

# =============== SMTP / Paths / Constantes ===============
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT   = 587
SMTP_TIMEOUT = 20  # timeout para conexiones SMTP
SERVICE_NAME = "ConfirmRecolectaSMTP"

APPDATA  = os.getenv("APPDATA") or os.path.expanduser("~")
DATA_DIR = os.path.join(APPDATA, "ConfirmRecolecta")
os.makedirs(DATA_DIR, exist_ok=True)
CONFIG   = os.path.join(DATA_DIR, "config.json")

INITIAL_CARRIERS = ["APOLOTRAN", "CHR", "ESGARI", "CEVA"]

# Sesión en memoria (no persistente)
SESSION = {"user": None, "pass": None}

# =============== Configuración persistente (no credenciales) ===============
def _default_cfg():
    return {
        "destinos": [],
        "cc_por_carrier": {c: [] for c in INITIAL_CARRIERS},
        "plantillas": {},        # {nombre: {subject, campos:{...}, comentarios, carrier, to, cc}}
        "freq_destinos": {},     # {email: count}
        "freq_carriers": {},     # {carrier: count}
        "smtp_user": None,       # legacy
        "smtp_pass": None        # legacy
    }

def load_config():
    if not os.path.exists(CONFIG):
        cfg = _default_cfg()
        save_config(cfg); return cfg
    try:
        with open(CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        cfg = _default_cfg()
    # migraciones defensivas
    base = _default_cfg()
    for k, v in base.items():
        if k not in cfg: cfg[k] = v
    for c in INITIAL_CARRIERS:
        cfg["cc_por_carrier"].setdefault(c, [])
    save_config(cfg); return cfg

def save_config(cfg):
    with open(CONFIG, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

cfg            = load_config()
destinos       = cfg["destinos"]
cc_por_carrier = cfg["cc_por_carrier"]
plantillas     = cfg["plantillas"]
freq_destinos  = cfg["freq_destinos"]
freq_carriers  = cfg["freq_carriers"]

# =============== Ranking para autocompletado ===============
def _sorted_by_usage(items, freq_map):
    return sorted(set(items), key=lambda x: (-int(freq_map.get(x, 0)), x.lower()))

def bump_freq(freq_map, key, inc=1):
    if not key: return
    freq_map[key] = int(freq_map.get(key, 0)) + inc
    save_config(cfg)

# =============== App (Tk) ===============
root = tk.Tk()
root.title("Elbrus OPS")
try:
    root.state('zoomed')
except Exception:
    root.geometry("1400x820")
root.minsize(1024, 680)

# =============== Paleta (CLARO) ===============
BG        = "#ffffff"
BG_CARD   = "#ffffff"
FG_TEXT   = "#111827"
FG_MUTED  = "#4b5563"
ACCENT    = "#1f6feb"
ACCENT_H  = "#185ec7"
DANGER    = "#dc2626"
DANGER_H  = "#b91c1c"
TAB_BG    = "#2a55b3"
TAB_SEL   = "#2e8b57"
TAB_FG    = "white"

root.configure(bg=BG)

# =============== Tema ttk + estilos ======================
style = ttk.Style(root)
style.theme_use("clam")
style.configure(".", background=BG, foreground=FG_TEXT, fieldbackground=BG_CARD)
style.configure('Custom.TNotebook', background=BG, borderwidth=0)
style.configure('Custom.TNotebook.Tab',
                background=TAB_BG, foreground=TAB_FG,
                padding=[14, 9], font=("Segoe UI", 10, "bold"))
style.map('Custom.TNotebook.Tab',
          background=[('selected', TAB_SEL), ('!selected', TAB_BG), ('active', TAB_BG), ('disabled', TAB_BG)],
          foreground=[('selected', TAB_FG), ('!selected', TAB_FG), ('active', TAB_FG), ('disabled', TAB_FG)])
style.configure("Card.TFrame", background=BG_CARD, relief="flat")
style.configure("CardInner.TFrame", background=BG_CARD)
style.configure("TLabel", background=BG_CARD, foreground=FG_TEXT, font=("Segoe UI", 10))
style.configure("Header.TLabel", background=BG, foreground=FG_TEXT, font=("Segoe UI", 18, "bold"))
style.configure("Subheader.TLabel", background=BG, foreground=FG_MUTED, font=("Segoe UI", 10))
style.configure("CardTitle.TLabel", background=BG_CARD, foreground=FG_TEXT, font=("Segoe UI", 12, "bold"))

# =============== Botones helper ===========================
def tk_button(parent, text, bg, fg, bg_active, command, bold=True, border=0):
    font = ("Segoe UI", 10, "bold") if bold else ("Segoe UI", 10)
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, activebackground=bg_active, activeforeground=fg,
                    bd=border, padx=14, pady=8, font=font,
                    relief="solid" if border else "flat", highlightthickness=0)
    return btn
def btn_accent(parent, text, command): return tk_button(parent, text, ACCENT, "white", ACCENT_H, command)
def btn_danger(parent, text, command): return tk_button(parent, text, DANGER, "white", DANGER_H, command)
def btn_plain(parent, text, command):  return tk_button(parent, text, BG_CARD, FG_TEXT, "#e5e7eb", command, bold=False, border=1)

# =============== Utilidades de UI / Hilos =================
status = tk.StringVar(value="Listo.")
def set_status(msg: str, timeout_ms: int = 4000):
    def _set():
        status.set(msg)
        if timeout_ms:
            root.after(timeout_ms, lambda: status.set("Listo."))
    root.after(0, _set)

def ui(fn, *args, **kwargs):
    """Invocar en hilo de UI."""
    root.after(0, lambda: fn(*args, **kwargs))

def call_on_ui(fn, *args, **kwargs):
    """
    Ejecuta fn en el hilo de UI y espera el resultado (para usar desde hilos).
    Devuelve lo que retorne fn.
    """
    done = threading.Event()
    out = {"result": None}
    def wrapper():
        out["result"] = fn(*args, **kwargs)
        done.set()
    root.after(0, wrapper)
    done.wait()
    return out["result"]

def make_card(parent, title: str):
    outer = ttk.Frame(parent, style="Card.TFrame")
    outer.pack(fill="both", expand=True, padx=14, pady=14)
    title_bar = ttk.Frame(outer, style="CardInner.TFrame"); title_bar.pack(fill="x", padx=14, pady=(14, 6))
    ttk.Label(title_bar, text=title, style="CardTitle.TLabel").pack(side="left")
    ttk.Separator(outer).pack(fill="x", padx=14, pady=(0, 10))
    inner = ttk.Frame(outer, style="CardInner.TFrame")
    inner.pack(fill="both", expand=True, padx=14, pady=(0, 14))
    inner.grid_columnconfigure(1, weight=1)
    return inner

# ====== Indicador de sesión en encabezado ======
session_user = tk.StringVar(value="Sin sesión")
def update_session_badge(user: str | None):
    if user:
        session_user.set(f"Sesión: {user}")
        session_badge.config(bg="#10b981", fg="white")
    else:
        session_user.set("Sin sesión")
        session_badge.config(bg="#e5e7eb", fg=FG_TEXT)

# =============== Credenciales / Login =====================
def get_stored_credentials():
    return SESSION["user"], SESSION["pass"]

def store_credentials(user, pwd):
    SESSION["user"], SESSION["pass"] = user, pwd
    if not ALLOW_PERSIST:
        cfg["smtp_user"] = None; cfg["smtp_pass"] = None; save_config(cfg)
        if keyring:
            try: keyring.delete_password(SERVICE_NAME, user)
            except Exception: pass
        return
    cfg["smtp_user"] = user
    if keyring:
        try:
            keyring.set_password(SERVICE_NAME, user, pwd)
            cfg["smtp_pass"] = None
        except Exception:
            cfg["smtp_pass"] = pwd
    else:
        cfg["smtp_pass"] = pwd
    save_config(cfg)

def clear_credentials():
    u = SESSION["user"]
    SESSION["user"] = None
    SESSION["pass"] = None
    cfg["smtp_user"] = None
    cfg["smtp_pass"] = None
    save_config(cfg)
    if u and keyring:
        try: keyring.delete_password(SERVICE_NAME, u)
        except Exception: pass

def test_credentials(user: str, pwd: str) -> bool:
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT) as s:
            s.starttls(); s.login(user, pwd)
        return True
    except Exception:
        return False

def ask_login(parent=None, prefill_user=""):
    dlg = tk.Toplevel(parent or root)
    dlg.title("Login SMTP"); dlg.transient(parent or root)
    dlg.grab_set(); dlg.resizable(False, False)

    frame = ttk.Frame(dlg, style="CardInner.TFrame"); frame.pack(fill="both", expand=True, padx=16, pady=16)
    ttk.Label(frame, text="Correo (usuario SMTP):").grid(row=0, column=0, sticky="e", padx=6, pady=6)
    e_user = ttk.Entry(frame, width=38); e_user.grid(row=0, column=1, sticky="we", padx=6, pady=6)
    if prefill_user: e_user.insert(0, prefill_user)
    ttk.Label(frame, text="Contraseña:").grid(row=1, column=0, sticky="e", padx=6, pady=6)
    e_pass = ttk.Entry(frame, width=38, show="*"); e_pass.grid(row=1, column=1, sticky="we", padx=6, pady=6)

    btns = ttk.Frame(frame, style="CardInner.TFrame"); btns.grid(row=2, column=0, columnspan=2, pady=(10,0), sticky="e")

    def on_ok():
        u, p = e_user.get().strip(), e_pass.get().strip()
        if not u or not p:
            messagebox.showwarning("Faltan datos", "Ingresa correo y contraseña."); return
        set_status("Probando credenciales…", 0)
        if test_credentials(u, p):
            store_credentials(u, p)
            update_session_badge(u)
            set_status("Sesión iniciada correctamente.")
            messagebox.showinfo("Login exitoso", f"Has iniciado sesión como: {u}")
            dlg.destroy()
        else:
            set_status("Error de autenticación.")
            messagebox.showerror("Autenticación fallida", "No fue posible iniciar sesión con esas credenciales.")
    def on_cancel():
        dlg.destroy()

    btn_plain(btns, "Cancelar", on_cancel).pack(side="right", padx=(6,0))
    btn_accent(btns, "Iniciar sesión", on_ok).pack(side="right", padx=6)

    dlg.wait_window()
    return SESSION["user"], SESSION["pass"]

def ensure_credentials(parent=None):
    user, pwd = get_stored_credentials()
    if not user or not pwd:
        user, pwd = ask_login(parent, prefill_user=user or "")
    return user, pwd

# =============== Envío SMTP (timeout + re-login en UI) ====
def smtp_send(msg: EmailMessage, parent=None):
    """
    Llamar desde hilo worker. Si necesita login, abre diálogo en UI con call_on_ui.
    """
    user, pwd = get_stored_credentials()
    if not user or not pwd:
        user, pwd = call_on_ui(ask_login, parent, user or "")
    if not user or not pwd:
        raise RuntimeError("No se ingresaron credenciales SMTP.")

    msg["From"] = user
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT) as s:
            s.starttls(); s.login(user, pwd); s.send_message(msg)
        return user
    except smtplib.SMTPAuthenticationError:
        call_on_ui(messagebox.showwarning, "Autenticación requerida",
                   "Tu sesión SMTP expiró o la contraseña cambió.\nIngresa credenciales de nuevo.")
        user2, pwd2 = call_on_ui(ask_login, parent, user or "")
        if not user2 or not pwd2:
            raise
        if "From" in msg:
            msg.replace_header("From", user2)
        else:
            msg["From"] = user2
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT) as s:
            s.starttls(); s.login(user2, pwd2); s.send_message(msg)
        call_on_ui(update_session_badge, user2)
        return user2

# =============== Encabezado ===============================
header = ttk.Frame(root, style="CardInner.TFrame"); header.pack(fill="x")
left_head = ttk.Frame(header, style="CardInner.TFrame"); left_head.pack(side="left", fill="x", expand=True, padx=10, pady=10)
ttk.Label(left_head, text="🚛 Elbrus OPS", style="Header.TLabel").pack(anchor="w")
ttk.Label(left_head, text="Confirmaciones, solicitudes y cotizaciones por correo", style="Subheader.TLabel").pack(anchor="w", pady=(2,0))
right_head = ttk.Frame(header, style="CardInner.TFrame"); right_head.pack(side="right", padx=10, pady=10)

session_badge = tk.Label(right_head, textvariable=session_user, padx=12, pady=6,
                         bg="#e5e7eb", fg=FG_TEXT, font=("Segoe UI", 10, "bold"))
session_badge.pack(side="left", padx=(0,8))
btn_accent(right_head, "Cuenta / Login", lambda: ask_login(root)).pack(side="left", padx=(0,8))

def do_logout():
    if SESSION["user"] is None:
        messagebox.showinfo("Cerrar sesión", "No hay una sesión activa."); return
    u = SESSION["user"]; clear_credentials(); update_session_badge(None)
    set_status("Sesión cerrada."); messagebox.showinfo("Cerrar sesión", f"Se cerró la sesión de: {u}")

btn_danger(right_head, "Cerrar sesión", do_logout).pack(side="left")
ttk.Separator(root, orient="horizontal").pack(fill="x", padx=10, pady=(0,8))
update_session_badge(None)

# =============== Contenedor principal (Notebook) ==========
container = ttk.Frame(root, style="CardInner.TFrame"); container.pack(fill="both", expand=True, padx=10, pady=(0,8))
notebook = ttk.Notebook(container, style='Custom.TNotebook'); notebook.pack(fill='both', expand=True)

# =============== Estado global ============================
last_carrier    = None
attachment_path = None

# =============== Pestaña 1: Confirmación ==================
tab_conf = ttk.Frame(notebook, style="CardInner.TFrame"); notebook.add(tab_conf, text="Confirmación")
frame_conf = make_card(tab_conf, "Datos de la confirmación")

ttk.Label(frame_conf, text="Referencia (Drive):").grid(row=0, column=0, sticky="e", padx=6, pady=6)
entry_ref = ttk.Entry(frame_conf, width=30); entry_ref.grid(row=0, column=1, sticky="we", padx=6, pady=6)

ttk.Label(frame_conf, text="Fecha (dd-mmm-aaaa):").grid(row=1, column=0, sticky="e", padx=6, pady=6)
entry_fecha = ttk.Entry(frame_conf, width=30); entry_fecha.grid(row=1, column=1, sticky="we", padx=6, pady=6)

ttk.Label(frame_conf, text="Destino (email):").grid(row=2, column=0, sticky="e", padx=6, pady=6)
combo_dest = ttk.Combobox(frame_conf, width=30, values=_sorted_by_usage(destinos, freq_destinos))
combo_dest.grid(row=2, column=1, sticky="we", padx=6, pady=6)
if destinos:
    combo_dest.set(_sorted_by_usage(destinos, freq_destinos)[0])

def refresh_destinos():
    combo_dest['values'] = _sorted_by_usage(destinos, freq_destinos)

def on_keyrelease_dest(event):
    txt = combo_dest.get().lower()
    base = _sorted_by_usage(destinos, freq_destinos)
    combo_dest['values'] = [d for d in base if txt in d.lower()] if txt else base
combo_dest.bind('<KeyRelease>', on_keyrelease_dest)

btn_send_conf = None

def send_confirmation():
    ref, fecha, dest = entry_ref.get().strip(), entry_fecha.get().strip(), combo_dest.get().strip()
    if not (ref and fecha and dest):
        messagebox.showwarning("Faltan datos", "Completa todos los campos."); return

    msg = EmailMessage()
    msg["To"] = dest
    msg["Subject"] = f"Confirmación de recolección: {ref}"
    body = ("Buen día,\n\n"
            "Comparto la información de la recolección:\n\n"
            f" • Fecha programada: {fecha}\n"
            f" • Referencia en Drive: {ref}\n\n"
            "Regreso con los datos de unidad.\n\nSaludos,")
    msg.set_content(body)

    def _do_send():
        try:
            set_status("Conectando a SMTP…", 0)
            sender_used = smtp_send(msg, parent=root)
            set_status(f"Confirmación enviada a {dest}")
            ui(messagebox.showinfo, "Enviado", f"Correo enviado a {dest}\nRemitente: {sender_used}")
            # Limpiezas y ranking
            ui(entry_ref.delete, 0, tk.END)
            ui(entry_fecha.delete, 0, tk.END)
            if dest not in destinos:
                destinos.append(dest); save_config(cfg)
            bump_freq(freq_destinos, dest, inc=1)
            ui(refresh_destinos)
        except RuntimeError as e:
            set_status("Envío cancelado.")
            ui(messagebox.showinfo, "Envío cancelado", str(e))
        except (smtplib.SMTPException, socket.timeout, OSError) as e:
            set_status("Error al enviar.")
            ui(messagebox.showerror, "Error al enviar", str(e))
        finally:
            if btn_send_conf: ui(btn_send_conf.config, state="normal")

    if btn_send_conf: btn_send_conf.config(state="disabled")
    threading.Thread(target=_do_send, daemon=True).start()

row_btn = ttk.Frame(frame_conf, style="CardInner.TFrame"); row_btn.grid(row=3, column=0, columnspan=2, sticky="e", pady=(8,0))
btn_send_conf = btn_accent(row_btn, "Enviar Confirmación", send_confirmation)
btn_send_conf.pack(side="right")

# =============== Pestaña 2: Solicitud / Cotización =========
tab_sol = ttk.Frame(notebook, style="CardInner.TFrame"); notebook.add(tab_sol, text="Solicitud / Cotización")
form_card = make_card(tab_sol, "Formulario de solicitud/cotización")

labels = [
    "Asunto del correo:", "Nombre de proveedor:", "Dirección de recolección:",
    "RFC:", "Contacto (Nombre y teléfono):", "Fecha y hora de recolección:",
    "Núm. de bultos (peso y dimensión):", "Edificio Jabil (entrega):", "Proyecto:"
]
entries = {}
for i, text in enumerate(labels):
    ttk.Label(form_card, text=text).grid(row=i, column=0, sticky="e", padx=6, pady=5)
    e = ttk.Entry(form_card, width=40); e.grid(row=i, column=1, sticky="we", padx=6, pady=5)
    entries[text] = e

entry_subject  = entries["Asunto del correo:"]
entry_prov     = entries["Nombre de proveedor:"]
entry_dir      = entries["Dirección de recolección:"]
entry_rfc      = entries["RFC:"]
entry_contacto = entries["Contacto (Nombre y teléfono):"]
entry_fecha2   = entries["Fecha y hora de recolección:"]
entry_bultos   = entries["Núm. de bultos (peso y dimensión):"]
entry_edif     = entries["Edificio Jabil (entrega):"]
entry_proj     = entries["Proyecto:"]

# --------- Plantillas ----------
tpl_card = make_card(tab_sol, "Plantillas")
ttk.Label(tpl_card, text="Plantilla:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
combo_tpl = ttk.Combobox(tpl_card, width=37, values=sorted(plantillas.keys()))
combo_tpl.grid(row=0, column=1, sticky="we", padx=6, pady=6)
tpl_btns = ttk.Frame(tpl_card, style="CardInner.TFrame"); tpl_btns.grid(row=0, column=2, padx=6, pady=6, sticky="w")

def _collect_form_data():
    return {
        "asunto": entry_subject.get().strip(),
        "proveedor": entry_prov.get().strip(),
        "direccion": entry_dir.get().strip(),
        "rfc": entry_rfc.get().strip(),
        "contacto": entry_contacto.get().strip(),
        "fecha_hora": entry_fecha2.get().strip(),
        "bultos": entry_bultos.get().strip(),
        "edificio": entry_edif.get().strip(),
        "proyecto": entry_proj.get().strip(),
        "comentarios": text_coment.get("1.0", tk.END).strip(),
        "carrier": combo_car.get().strip(),
        "to": entry_to.get().strip(),
        "cc": entry_cc.get().strip()
    }

def _fill_form_from_data(d):
    entry_subject.delete(0, tk.END); entry_subject.insert(0, d.get("asunto",""))
    entry_prov.delete(0, tk.END);    entry_prov.insert(0, d.get("proveedor",""))
    entry_dir.delete(0, tk.END);     entry_dir.insert(0, d.get("direccion",""))
    entry_rfc.delete(0, tk.END);     entry_rfc.insert(0, d.get("rfc",""))
    entry_contacto.delete(0, tk.END);entry_contacto.insert(0, d.get("contacto",""))
    entry_fecha2.delete(0, tk.END);  entry_fecha2.insert(0, d.get("fecha_hora",""))
    entry_bultos.delete(0, tk.END);  entry_bultos.insert(0, d.get("bultos",""))
    entry_edif.delete(0, tk.END);    entry_edif.insert(0, d.get("edificio",""))
    entry_proj.delete(0, tk.END);    entry_proj.insert(0, d.get("proyecto",""))
    text_coment.delete("1.0", tk.END); text_coment.insert("1.0", d.get("comentarios",""))
    if d.get("carrier"): combo_car.set(d["carrier"]); on_carrier_change()
    if d.get("to") is not None: entry_to.delete(0, tk.END); entry_to.insert(0, d.get("to",""))
    if d.get("cc") is not None: entry_cc.delete(0, tk.END); entry_cc.insert(0, d.get("cc",""))

def save_template():
    name = combo_tpl.get().strip()
    if not name:
        messagebox.showwarning("Plantillas", "Escribe un nombre en el campo 'Plantilla' para guardar."); return
    data = _collect_form_data()
    plantillas[name] = {
        "subject": data["asunto"],
        "campos": {
            "proveedor": data["proveedor"], "direccion": data["direccion"], "rfc": data["rfc"],
            "contacto": data["contacto"], "fecha_hora": data["fecha_hora"], "bultos": data["bultos"],
            "edificio": data["edificio"], "proyecto": data["proyecto"]
        },
        "comentarios": data["comentarios"],
        "carrier": data["carrier"],
        "to": data["to"],
        "cc": data["cc"]
    }
    save_config(cfg)
    combo_tpl['values'] = sorted(plantillas.keys())
    set_status(f"Plantilla '{name}' guardada.")

def apply_template():
    name = combo_tpl.get().strip()
    if not name or name not in plantillas:
        messagebox.showwarning("Plantillas", "Selecciona una plantilla existente para aplicar."); return
    t = plantillas[name]
    d = {
        "asunto": t.get("subject",""),
        "proveedor": t["campos"].get("proveedor",""),
        "direccion": t["campos"].get("direccion",""),
        "rfc": t["campos"].get("rfc",""),
        "contacto": t["campos"].get("contacto",""),
        "fecha_hora": t["campos"].get("fecha_hora",""),
        "bultos": t["campos"].get("bultos",""),
        "edificio": t["campos"].get("edificio",""),
        "proyecto": t["campos"].get("proyecto",""),
        "comentarios": t.get("comentarios",""),
        "carrier": t.get("carrier",""),
        "to": t.get("to",""),
        "cc": t.get("cc",""),
    }
    _fill_form_from_data(d)
    set_status(f"Plantilla '{name}' aplicada.")

def delete_template():
    name = combo_tpl.get().strip()
    if not name or name not in plantillas:
        messagebox.showwarning("Plantillas", "Selecciona una plantilla existente para eliminar."); return
    if messagebox.askyesno("Eliminar plantilla", f"¿Eliminar '{name}'?"):
        del plantillas[name]; save_config(cfg)
        combo_tpl['values'] = sorted(plantillas.keys()); combo_tpl.set("")
        set_status(f"Plantilla '{name}' eliminada.")

btn_plain(tpl_btns, "Aplicar", apply_template).pack(side="left", padx=(0,6))
btn_accent(tpl_btns, "Guardar como", save_template).pack(side="left", padx=(0,6))
btn_danger(tpl_btns, "Eliminar", delete_template).pack(side="left")

# --------- Comentarios y adjuntos ----------
coment_card = make_card(tab_sol, "Comentarios y adjuntos")
ttk.Label(coment_card, text="Comentarios adicionales:").grid(row=0, column=0, sticky="ne", padx=6, pady=5)
text_coment = tk.Text(coment_card, width=40, height=5, bd=1, relief="solid")
text_coment.configure(bg="#ffffff", fg=FG_TEXT, insertbackground=FG_TEXT,
                      highlightbackground="#d1d5db", highlightcolor="#3b82f6")
text_coment.grid(row=0, column=1, sticky="we", padx=6, pady=5)

# --------- Carrier y CC (ranking) ----------
ttk.Label(form_card, text="Para (To):").grid(row=len(labels), column=0, sticky="e", padx=6, pady=5)
entry_to = ttk.Entry(form_card, width=40); entry_to.grid(row=len(labels), column=1, sticky="we", padx=6, pady=5)

ttk.Label(form_card, text="Carrier:").grid(row=len(labels)+1, column=0, sticky="e", padx=6, pady=5)
combo_car = ttk.Combobox(form_card, width=37, values=_sorted_by_usage(list(cc_por_carrier.keys()), freq_carriers), state="normal")
combo_car.grid(row=len(labels)+1, column=1, sticky="we", padx=6, pady=5)

ttk.Label(form_card, text="CC (separa con comas):").grid(row=len(labels)+2, column=0, sticky="e", padx=6, pady=5)
entry_cc = ttk.Entry(form_card, width=40); entry_cc.grid(row=len(labels)+2, column=1, sticky="we", padx=6, pady=5)

last_carrier = None
def _refresh_carriers_combobox():
    combo_car['values'] = _sorted_by_usage(list(cc_por_carrier.keys()), freq_carriers)

def on_carrier_change(event=None):
    """
    - Guarda los CC del carrier anterior.
    - Carga CC del carrier actual si existe.
    - NO crea carriers nuevos automáticamente (evita recrear tras eliminar).
    """
    global last_carrier
    c = combo_car.get().strip()
    # guardar CC del previo
    if last_carrier and last_carrier in cc_por_carrier:
        cc_por_carrier[last_carrier] = [x.strip() for x in entry_cc.get().split(",") if x.strip()]
        save_config(cfg)

    if not c:
        entry_cc.delete(0, tk.END)
        last_carrier = None
        return

    if c in cc_por_carrier:
        entry_cc.delete(0, tk.END)
        entry_cc.insert(0, ", ".join(cc_por_carrier[c]))
        last_carrier = c
    else:
        # No existe: no lo creamos
        entry_cc.delete(0, tk.END)
        last_carrier = None

combo_car.bind("<<ComboboxSelected>>", on_carrier_change)
combo_car.bind("<FocusOut>", on_carrier_change)

def add_carrier():
    """
    Diálogo para agregar un nuevo carrier (con CC inicial opcional).
    """
    dlg = tk.Toplevel(root)
    dlg.title("Agregar Carrier"); dlg.transient(root); dlg.grab_set(); dlg.resizable(False, False)

    frame = ttk.Frame(dlg, style="CardInner.TFrame"); frame.pack(fill="both", expand=True, padx=16, pady=16)
    ttk.Label(frame, text="Nombre del carrier:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
    e_name = ttk.Entry(frame, width=40); e_name.grid(row=0, column=1, sticky="we", padx=6, pady=6)

    ttk.Label(frame, text="CC inicial (opcional, coma-separado):").grid(row=1, column=0, sticky="e", padx=6, pady=6)
    e_cc = ttk.Entry(frame, width=40); e_cc.grid(row=1, column=1, sticky="we", padx=6, pady=6)

    btns = ttk.Frame(frame, style="CardInner.TFrame"); btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(10,0))

    def on_ok():
        name = e_name.get().strip()
        if not name:
            messagebox.showwarning("Carrier", "Escribe un nombre."); return
        # Normaliza: mayúsculas (coherencia con iniciales) y sin dobles espacios
        name_norm = " ".join(name.upper().split())
        if name_norm in cc_por_carrier:
            messagebox.showinfo("Carrier", f"'{name_norm}' ya existe."); return
        cc_init = [x.strip() for x in e_cc.get().split(",") if x.strip()]
        cc_por_carrier[name_norm] = cc_init
        save_config(cfg)
        # Refrescar y seleccionar
        _refresh_carriers_combobox()
        combo_car.set(name_norm)
        entry_cc.delete(0, tk.END)
        entry_cc.insert(0, ", ".join(cc_init))
        global last_carrier
        last_carrier = name_norm
        set_status(f"Carrier '{name_norm}' agregado.", 3000)
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    btn_plain(btns, "Cancelar", on_cancel).pack(side="right", padx=(6,0))
    btn_accent(btns, "Agregar", on_ok).pack(side="right", padx=6)

    # Enter para aceptar
    e_name.focus_set()
    dlg.bind("<Return>", lambda _e: on_ok())
    dlg.wait_window()

def delete_carrier():
    """
    Elimina el carrier seleccionado y limpia estados para que NO se recree.
    """
    global last_carrier
    c = combo_car.get().strip()
    if not c or c not in cc_por_carrier:
        messagebox.showinfo("Eliminar Carrier", "Selecciona un carrier válido.")
        return
    if messagebox.askyesno("Eliminar", f"¿Eliminar carrier '{c}'?"):
        del cc_por_carrier[c]
        save_config(cfg)
        last_carrier = None
        combo_car.set("")
        entry_cc.delete(0, tk.END)
        _refresh_carriers_combobox()
        set_status(f"Carrier '{c}' eliminado.", 3000)

def attach_file():
    global attachment_path
    p = filedialog.askopenfilename(title="Selecciona archivo a adjuntar",
                                   filetypes=[("Todos los archivos","*.*")])
    if p:
        attachment_path = p; lbl_att.configure(text=os.path.basename(p))
    else:
        attachment_path = None; lbl_att.configure(text="(ningún archivo)")

actions = ttk.Frame(tab_sol, style="CardInner.TFrame"); actions.pack(fill="x", padx=24, pady=(0,12))
btn_plain(actions, "Agregar Carrier", add_carrier).pack(side="left")  # <-- NUEVO
btn_danger(actions, "Eliminar Carrier", delete_carrier).pack(side="left", padx=(8,0))
btn_plain(actions, "Adjuntar Archivo", attach_file).pack(side="left", padx=(8,0))
lbl_att = ttk.Label(actions, text="(ningún archivo)"); lbl_att.pack(side="left", padx=8)

var_det = tk.BooleanVar(value=True)
ttk.Checkbutton(actions, text="Incluir detalles en cuerpo", variable=var_det).pack(side="right")

# =============== Render de variables en plantillas ===============
VAR_PATTERN = re.compile(r"\{([a-zA-Z0-9_]+)\}")
def _entries_to_varmap():
    return {
        "proveedor": entry_prov.get().strip(),
        "direccion": entry_dir.get().strip(),
        "rfc": entry_rfc.get().strip(),
        "contacto": entry_contacto.get().strip(),
        "fecha_hora": entry_fecha2.get().strip(),
        "bultos": entry_bultos.get().strip(),
        "edificio": entry_edif.get().strip(),
        "proyecto": entry_proj.get().strip(),
        "carrier": combo_car.get().strip(),
        "to": entry_to.get().strip(),
        "cc": entry_cc.get().strip()
    }
def render_vars(text: str, varmap: dict):
    def repl(m):
        key = m.group(1).lower()
        return varmap.get(key, m.group(0))
    return VAR_PATTERN.sub(repl, text or "")

# =============== Envío de Solicitud/Cotización (asíncrono) =======
btn_send_sol = None
btn_send_cot = None

def send_solicitud(template):
    global attachment_path, last_carrier
    data        = {k: v.get().strip() for k, v in entries.items()}
    comentarios = text_coment.get("1.0", tk.END).strip()
    carrier     = combo_car.get().strip()
    to_addr     = entry_to.get().strip()
    cc_list     = [x.strip() for x in entry_cc.get().split(",") if x.strip()]
    subject     = data["Asunto del correo:"]

    if not (subject and carrier and to_addr):
        messagebox.showwarning("Faltan datos","Asunto, Para y Carrier son obligatorios"); return

    varmap = _entries_to_varmap()
    subject = render_vars(subject, varmap)
    comentarios = render_vars(comentarios, varmap)
    plain_rows = [(lbl, render_vars(data.get(lbl,""), varmap)) for lbl in labels[1:]]

    if template == "cotización":
        intro = "Buen día equipo,\n\nMe apoyarían revisando disponibilidad y cotización, por favor.\n\n"
        outro = "\nQuedo atento.\n\nSaludos,"
    else:
        intro = "Buen día equipo,\n\nMe apoyan con la siguiente recolección, por favor.\n\n"
        outro = "\nQuedo atento a su confirmación y datos de unidad.\n\nSaludos,"

    if var_det.get():
        rows_plain = plain_rows[:]
        if comentarios: rows_plain.append(("Comentarios adicionales", comentarios))
        rows_plain.append(("Carrier", carrier))
        plain = intro + "\n".join(f"{l}: {v}" for l, v in rows_plain) + outro
        rows_html = [(l, html.escape(v)) for l, v in rows_plain]
        html_body = f"<html><body><p>{intro.replace(chr(10),'<br>')}</p><table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse; font-family:Segoe UI, Arial, sans-serif; font-size:13px;'>"
        for l, v in rows_html:
            html_body += f"<tr><th align='left' style='background:#f3f4f6; color:{FG_TEXT}; padding:6px;'>{l}</th><td style='background:#ffffff; color:{FG_TEXT}; padding:6px;'>{v}</td></tr>"
        html_body += f"</table><p>{outro.replace(chr(10),'<br>')}</p></body></html>"
    else:
        plain = intro + outro
        html_body  = f"<html><body><p>{intro.replace(chr(10),'<br>')}</p><p>{outro.replace(chr(10),'<br>')}</p></body></html>"

    msg = EmailMessage()
    msg["To"] = to_addr
    if cc_list: msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.set_content(plain); msg.add_alternative(html_body, subtype="html")

    if attachment_path:
        mime_type, _ = mimetypes.guess_type(attachment_path)
        maintype, subtype = ('application','octet-stream') if not mime_type else mime_type.split('/',1)
        with open(attachment_path,'rb') as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    def _do_send():
        try:
            set_status("Enviando correo…", 0)
            sender_used = smtp_send(msg, parent=root)
            set_status(f"{template.capitalize()} enviada")
            ui(messagebox.showinfo, "Enviado", f"{template.capitalize()} enviada.\nRemitente: {sender_used}")
            # Limpieza parcial
            ui(entry_subject.delete, 0, tk.END)
            ui(entry_to.delete, 0, tk.END)
            for w in (entries["Nombre de proveedor:"], entries["Dirección de recolección:"],
                      entries["RFC:"], entries["Contacto (Nombre y teléfono):"],
                      entries["Fecha y hora de recolección:"], entries["Núm. de bultos (peso y dimensión):"],
                      entries["Edificio Jabil (entrega):"], entries["Proyecto:"]):
                ui(w.delete, 0, tk.END)
            ui(text_coment.delete, "1.0", tk.END)

            def _reset_att():
                global attachment_path
                attachment_path = None
                lbl_att.config(text="(ningún archivo)")
            ui(_reset_att)

            # Persistir y ranking
            if carrier in cc_por_carrier:
                cc_por_carrier[carrier] = cc_list
                save_config(cfg)
            bump_freq(freq_carriers, carrier, inc=1)
            ui(_refresh_carriers_combobox)
        except RuntimeError as e:
            set_status("Envío cancelado.")
            ui(messagebox.showinfo, "Envío cancelado", str(e))
        except (smtplib.SMTPException, socket.timeout, OSError) as e:
            set_status("Error al enviar.")
            ui(messagebox.showerror, "Error al enviar", str(e))
        finally:
            if template == "cotización" and btn_send_cot:
                ui(btn_send_cot.config, state="normal")
            if template == "solicitud" and btn_send_sol:
                ui(btn_send_sol.config, state="normal")

    if template == "cotización" and btn_send_cot: btn_send_cot.config(state="disabled")
    if template == "solicitud" and btn_send_sol: btn_send_sol.config(state="disabled")
    threading.Thread(target=_do_send, daemon=True).start()

send_row = ttk.Frame(tab_sol, style="CardInner.TFrame"); send_row.pack(fill="x", padx=24, pady=(0,14))
btn_send_sol = btn_plain(send_row, "Enviar Solicitud Recolección", lambda: send_solicitud("solicitud"))
btn_send_sol.pack(side="right")
btn_send_cot = btn_accent(send_row, "Enviar Cotización", lambda: send_solicitud("cotización"))
btn_send_cot.pack(side="right", padx=(0,8))

# =============== Barra de estado ==========================
status_bar = ttk.Frame(root, style="CardInner.TFrame"); status_bar.pack(fill="x", padx=10, pady=(6,8))
ttk.Label(status_bar, textvariable=status, style="Subheader.TLabel").pack(side="left")

# =============== Mainloop ================================
if __name__ == "__main__":
    root.mainloop()

