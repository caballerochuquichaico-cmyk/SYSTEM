# -*- coding: utf-8 -*-
"""
IAPAY S.R.L. - Atención al Comercio (SoundBox Q191 Aisino)
Aplicación de escritorio para Windows. Base de datos operativa: IAPAY_Atencion_Comercio.xlsx
(en la misma carpeta del .exe, o en la ruta indicada en config.json -> {"excel_path": "C:\\ruta\\archivo.xlsx"}).
"""
import json
import os
import re
import sys
import tkinter as tk
from collections import Counter
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ----------------------------------------------------------------------------
# Configuración y constantes
# ----------------------------------------------------------------------------
APP_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
FMT = "%d/%m/%Y %H:%M"
FMT_D = "%d/%m/%Y"
META_RESPUESTA_MIN = 15
META_PRIMER_CONTACTO = 70.0

SHEETS = {
    "Registro_Casos": ["Fecha_Hora_Contacto", "Hora_Primera_Respuesta", "Comercio", "Nro_Serie", "Canal_Tipo",
                       "Motivo", "Detalle_Solucion", "Estado", "Area_Derivacion", "Fecha_Cierre", "Minutos_Respuesta"],
    "Comercios_Bienvenidas": ["Comercio", "Telefono_WhatsApp", "Nro_Serie", "Ejecutivo_Campo", "Fecha_Instalacion",
                              "Bienvenida_Realizada", "Observaciones_Bienvenida"],
    "Activacion_Prospectos": ["Comercio_Prospecto", "Telefono", "Tipo_Gestion", "Cobra_Con_QR", "Cuenta_Banco_Economico",
                              "Motivo_Resultado", "Requiere_Visita_Operaciones", "Fecha_Agendada"],
}
KPI_SHEET = "KPIs_Indicadores"

ESTADOS = ["Cerrado en primer contacto", "Derivado a Operaciones"]
AREAS = ["", "Operaciones", "Finanzas y Administración", "Gerencia General"]
CANAL_TIPO = [f"{c} - {t}" for c in ("WhatsApp", "Llamada", "Web")
              for t in ("Soporte", "Reclamo", "Bienvenida", "Reactivación", "Activación")]
MOTIVOS = ["No habla", "Sin conexión", "No anunció un pago", "Consulta de comodato / condiciones",
           "Factura / comisión", "Solicitud de visita técnica", "Cambio de equipo", "Pérdida / robo del equipo",
           "Bienvenida", "Reactivación", "Otro"]

# ----------------------------------------------------------------------------
# Guardrails de comunicación
# ----------------------------------------------------------------------------
_DIAS = r"(hoy|ma[ñn]ana|lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo|\d{1,2}[/-]\d{1,2}|\d{1,2}:\d{2}|a\s+las\s+\d)"
RISKY = [
    (r"seguridad\s+garantizada", "«Seguridad garantizada» no se debe prometer."),
    (r"pagas?\s+solo\s+cuando\s+vendes|si\s+no\s+vendes\s+no\s+pagas", "No decir «pagas solo cuando vendes / si no vendes no pagas»."),
    (r"vendis", "No hablar mal de la competencia (Vendis) ni compararse."),
    (r"robos?\s+(del|de\s+los)\s+(personal|empleados|cajeros)|evit\w+\s+robos", "No ofrecer el equipo como solución para evitar robos del personal."),
    (r"descuento|rebaja|promoci[oó]n\s+especial", "No ofrecer descuentos ni condiciones especiales (Gerencia General decide)."),
    (r"devoluci[oó]n|devolver|reembols", "No ofrecer devoluciones."),
    (r"dar(emos)?\s+de\s+baja|dimos\s+de\s+baja|baja\s+del\s+comercio|cancelar(emos)?\s+(su\s+)?afiliaci", "No dar de baja comercios desde Atención al Comercio."),
    (r"(visit\w*|iremos|ir[aá]\s+(un\s+)?t[eé]cnico|t[eé]cnico\s+ir[aá]|cambi\w+|reemplaz\w+)[^.\n]{0,60}" + _DIAS,
     "No prometer fechas exactas de visita o cambio de equipo: solo Operaciones agenda."),
]


def check_text(text):
    low = (text or "").lower()
    return [msg for pat, msg in RISKY if re.search(pat, low)]


def guard_ok(data):
    issues = check_text(" ".join(v for v in data.values() if isinstance(v, str)))
    if not issues:
        return True
    return messagebox.askyesno("⚠️ Frase prohibida o riesgosa",
                               "Se detectó lo siguiente:\n\n• " + "\n• ".join(issues) +
                               "\n\n¿Guardar de todas formas?")


# ----------------------------------------------------------------------------
# Excel (base de datos)
# ----------------------------------------------------------------------------
def excel_path():
    cfg = os.path.join(APP_DIR, "config.json")
    if os.path.exists(cfg):
        try:
            with open(cfg, encoding="utf-8") as f:
                p = json.load(f).get("excel_path")
            if p:
                return p
        except Exception:
            pass
    return os.path.join(APP_DIR, "IAPAY_Atencion_Comercio.xlsx")


XLSX = excel_path()


def _style_header(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, c in enumerate(ws[1], 1):
        ws.column_dimensions[c.column_letter].width = max(14, len(str(c.value)) + 4)
    ws.freeze_panes = "A2"


def _build_kpi_sheet(ws):
    ws.append(["Indicador", "Valor", "Meta", "Estado"])
    R = "Registro_Casos"
    rows = [
        ("Total de casos atendidos", f"=COUNTA({R}!A2:A100000)", "", ""),
        ("Tiempo promedio de primera respuesta (min)", f"=IFERROR(AVERAGE({R}!K2:K100000),0)", META_RESPUESTA_MIN,
         '=IF(B3<=C3,"✅ Cumple","⚠️ Fuera de meta")'),
        ("% resolución en primer contacto", f'=IFERROR(COUNTIF({R}!H2:H100000,"Cerrado en primer contacto")/B2,0)',
         META_PRIMER_CONTACTO / 100, '=IF(B4>=C4,"✅ Cumple","⚠️ Fuera de meta")'),
        ("Casos derivados a Operaciones", f'=COUNTIF({R}!I2:I100000,"Operaciones")', "", ""),
        ("Casos derivados a Finanzas y Administración", f'=COUNTIF({R}!I2:I100000,"Finanzas y Administración")', "", ""),
        ("Casos derivados a Gerencia General", f'=COUNTIF({R}!I2:I100000,"Gerencia General")', "", ""),
        ("Bienvenidas pendientes", '=COUNTIF(Comercios_Bienvenidas!F2:F100000,"No")', "", ""),
    ]
    for r in rows:
        ws.append([v if v != "" else None for v in r])
    ws["B4"].number_format = "0.0%"
    ws["C4"].number_format = "0%"
    ws["B3"].number_format = "0.0"
    _style_header(ws)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["D"].width = 20


def ensure_workbook():
    if os.path.exists(XLSX):
        wb = load_workbook(XLSX)
        changed = False
        for name, cols in SHEETS.items():
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                ws.append(cols)
                _style_header(ws)
                changed = True
        if KPI_SHEET not in wb.sheetnames:
            _build_kpi_sheet(wb.create_sheet(KPI_SHEET))
            changed = True
        if changed:
            wb.save(XLSX)
        return
    wb = Workbook()
    wb.remove(wb.active)
    for name, cols in SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        _style_header(ws)
    _build_kpi_sheet(wb.create_sheet(KPI_SHEET))
    wb.save(XLSX)


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime(FMT)
    return str(v)


def read_rows(sheet):
    ensure_workbook()
    wb = load_workbook(XLSX, data_only=True)
    ws = wb[sheet]
    cols = SHEETS[sheet]
    out = []
    for i, row in enumerate(ws.iter_rows(min_row=2, max_col=len(cols), values_only=True), start=2):
        if all(v is None or v == "" for v in row):
            continue
        out.append((i, dict(zip(cols, [_cell_to_str(v) for v in row]))))
    return out


def _with_wb(fn):
    try:
        ensure_workbook()
        wb = load_workbook(XLSX)
        fn(wb)
        wb.save(XLSX)
        return True
    except PermissionError:
        messagebox.showerror("Excel en uso", "No se pudo guardar. Cierre el archivo\n" + XLSX +
                             "\nsi lo tiene abierto en Excel e intente de nuevo.")
    except Exception as e:  # noqa
        messagebox.showerror("Error al guardar en Excel", str(e))
    return False


def _val(v, col=""):
    if v == "":
        return None
    if col == "Minutos_Respuesta":
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None
    return v


def append_row(sheet, d):
    cols = SHEETS[sheet]
    return _with_wb(lambda wb: wb[sheet].append([_val(d.get(c, ""), c) for c in cols]))


def update_row(sheet, idx, d):
    cols = SHEETS[sheet]

    def fn(wb):
        ws = wb[sheet]
        for j, c in enumerate(cols, 1):
            ws.cell(row=idx, column=j, value=_val(d.get(c, ""), c))
    return _with_wb(fn)


# ----------------------------------------------------------------------------
# Utilidades de fecha
# ----------------------------------------------------------------------------
def now_s():
    return datetime.now().strftime(FMT)


def parse_dt(s):
    s = (s or "").strip()
    if not s:
        return None
    for f in (FMT, "%d/%m/%Y %H:%M:%S", FMT_D, "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def parse_resp(contact_dt, s):
    """Acepta fecha completa o solo hora (hh:mm, se toma la fecha del contacto)."""
    s = (s or "").strip()
    if not s:
        return None
    if re.fullmatch(r"\d{1,2}:\d{2}", s) and contact_dt:
        h, m = map(int, s.split(":"))
        return contact_dt.replace(hour=h, minute=m, second=0, microsecond=0)
    return parse_dt(s)


# ----------------------------------------------------------------------------
# Guías rápidas de soporte (copiables con un clic)
# ----------------------------------------------------------------------------
ALERTA_PAGO = "⚠️ Si el pago NO figura en plataforma, por favor *no entregue la mercadería* hasta que lo confirmemos."
GUIDES = {
    "1. Pedir datos (primero)": (
        "Hola{c} 👋 Gracias por escribir a IAPAY. Para ayudarle más rápido, ¿me indica por favor el "
        "*nombre de su comercio* y el *número de serie* de su SoundBox? 🙏"),
    "2. No habla": (
        "Hola{c} 👋 Revisemos juntos su SoundBox{s}:\n"
        "1️⃣ Verifique que esté *encendido*.\n"
        "2️⃣ *Suba el volumen*.\n"
        "3️⃣ Revise que tenga *batería* o conéctelo al *cargador*.\n"
        "Cuénteme qué pasó y le sigo ayudando 🙌"),
    "3. Sin conexión": (
        "Hola{c} 👋 Para recuperar la conexión de su SoundBox{s}:\n"
        "1️⃣ *Apáguelo y vuelva a encenderlo*.\n"
        "2️⃣ Revise que tenga *señal* (chip 4G Entel) o que esté conectado al *WiFi*.\n"
        "3️⃣ Mientras tanto, nosotros validamos su equipo en el panel de la plataforma.\n"
        "Avíseme si ya se conectó 🙌"),
    "4. No anunció un pago": (
        "Hola{c} 👋 Para verificar el pago, ¿me indica por favor el *monto* y la *hora aproximada*? "
        "Lo validamos en plataforma y le confirmo enseguida.\n" + ALERTA_PAGO),
    "5. Comodato y condiciones": (
        "Hola{c} 👋 Le recordamos las condiciones del SoundBox:\n"
        "• Se entrega en *comodato, sin costo*.\n"
        "• Comisión: *1%* por cobro.\n"
        "• Pérdida del equipo: *Bs 500*.\n"
        "• Robo: *Bs 250*, con denuncia policial dentro de *72 horas*.\n"
        "• Retiro del equipo si pasan *3 meses bajo el mínimo*.\n"
        "• Suspensión si pasan *60 días sin uso*.\n"
        "Cualquier duda, escríbanos por aquí 🙌"),
}


# ----------------------------------------------------------------------------
# Pestaña genérica de datos <-> Excel
# ----------------------------------------------------------------------------
class DataTab(ttk.Frame):
    sheet = ""
    fields = []      # (clave, etiqueta, tipo[entry|combo|ro|text], valores)
    show_cols = []

    def __init__(self, app, parent):
        super().__init__(parent)
        self.app = app
        self.cols = SHEETS[self.sheet]
        self.vars, self.texts, self.rows, self.editing = {}, {}, {}, None

        form = ttk.LabelFrame(self, text="Datos")
        form.pack(fill="x", padx=8, pady=(8, 4))
        r = c = 0
        for key, label, kind, vals in self.fields:
            if kind == "text":
                if c:
                    r, c = r + 1, 0
                ttk.Label(form, text=label).grid(row=r, column=0, sticky="ne", padx=4, pady=3)
                t = tk.Text(form, height=3, width=60, wrap="word")
                t.grid(row=r, column=1, columnspan=3, sticky="we", padx=4, pady=3)
                self.texts[key] = t
                r += 1
                continue
            ttk.Label(form, text=label).grid(row=r, column=c * 2, sticky="e", padx=4, pady=3)
            v = tk.StringVar()
            self.vars[key] = v
            if kind in ("combo", "ro"):
                w = ttk.Combobox(form, textvariable=v, values=vals, width=34,
                                 state="readonly" if kind == "ro" else "normal")
            else:
                w = ttk.Entry(form, textvariable=v, width=36)
            w.grid(row=r, column=c * 2 + 1, sticky="we", padx=4, pady=3)
            c += 1
            if c == 2:
                r, c = r + 1, 0
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        bf = ttk.Frame(self)
        bf.pack(fill="x", padx=8, pady=2)
        ttk.Button(bf, text="💾 Guardar como nuevo", command=lambda: self.save(False)).pack(side="left", padx=2)
        ttk.Button(bf, text="✏️ Guardar cambios", command=lambda: self.save(True)).pack(side="left", padx=2)
        ttk.Button(bf, text="🧹 Limpiar", command=self.clear).pack(side="left", padx=2)
        self.extra_buttons(bf)

        tf = ttk.Frame(self)
        tf.pack(fill="both", expand=True, padx=8, pady=6)
        self.tree = ttk.Treeview(tf, columns=self.show_cols, show="headings", height=9)
        for col in self.show_cols:
            self.tree.heading(col, text=col.replace("_", " "))
            self.tree.column(col, width=110, stretch=True)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.tag_configure("pend", background="#FFF3CD")
        self.clear()

    # ---- ganchos para subclases
    def defaults(self):
        return {}

    def extra_buttons(self, bf):
        pass

    def row_tag(self, row):
        return ()

    def validate(self, d):
        return None

    def before_save(self, full, old):
        return full

    def after_save(self, full, is_new):
        pass

    # ---- lógica común
    def get(self, key):
        if key in self.texts:
            return self.texts[key].get("1.0", "end").strip()
        return self.vars[key].get().strip()

    def set(self, key, val):
        if key in self.texts:
            self.texts[key].delete("1.0", "end")
            self.texts[key].insert("1.0", val)
        else:
            self.vars[key].set(val)

    def collect(self):
        return {k: self.get(k) for k in list(self.vars) + list(self.texts)}

    def clear(self):
        self.editing = None
        for v in self.vars.values():
            v.set("")
        for t in self.texts.values():
            t.delete("1.0", "end")
        for k, v in self.defaults().items():
            self.set(k, v)

    def refresh(self):
        self.rows = {}
        self.tree.delete(*self.tree.get_children())
        try:
            data = read_rows(self.sheet)
        except Exception as e:  # noqa
            self.app.status.set(f"No se pudo leer el Excel: {e}")
            return
        for idx, row in reversed(data):
            self.rows[str(idx)] = row
            self.tree.insert("", "end", iid=str(idx), values=[row.get(c, "") for c in self.show_cols],
                             tags=self.row_tag(row))

    def on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.rows[sel[0]]
        self.editing = int(sel[0])
        for k in self.vars:
            self.vars[k].set(row.get(k, ""))
        for k in self.texts:
            self.set(k, row.get(k, ""))

    def save(self, as_edit):
        if as_edit and self.editing is None:
            messagebox.showinfo("Seleccione un registro", "Elija una fila de la tabla para editarla.")
            return
        d = self.collect()
        err = self.validate(d)
        if err:
            messagebox.showwarning("Revise los datos", err)
            return
        if not guard_ok(d):
            return
        old = self.rows.get(str(self.editing), {}) if as_edit else {}
        full = self.before_save({**old, **d}, old)
        ok = update_row(self.sheet, self.editing, full) if as_edit else append_row(self.sheet, full)
        if not ok:
            return
        self.app.status.set(f"✅ Guardado en Excel ({now_s()})")
        self.after_save(full, not as_edit)
        self.clear()
        self.app.refresh_all()


# ----------------------------------------------------------------------------
# Registro de casos
# ----------------------------------------------------------------------------
class CasosTab(DataTab):
    sheet = "Registro_Casos"
    fields = [
        ("Fecha_Hora_Contacto", "Fecha y hora del contacto", "entry", None),
        ("Hora_Primera_Respuesta", "1ª respuesta (hh:mm o fecha hora)", "entry", None),
        ("Comercio", "Comercio *", "entry", None),
        ("Nro_Serie", "N° de serie *", "entry", None),
        ("Canal_Tipo", "Canal - Tipo", "combo", CANAL_TIPO),
        ("Motivo", "Motivo", "combo", MOTIVOS),
        ("Estado", "Estado", "ro", ESTADOS),
        ("Area_Derivacion", "Área de derivación", "ro", AREAS),
        ("Fecha_Cierre", "Fecha de cierre", "entry", None),
        ("Detalle_Solucion", "Detalle / solución", "text", None),
    ]
    show_cols = ["Fecha_Hora_Contacto", "Comercio", "Nro_Serie", "Canal_Tipo", "Motivo", "Estado",
                 "Area_Derivacion", "Minutos_Respuesta"]

    def defaults(self):
        return {"Fecha_Hora_Contacto": now_s(), "Estado": ESTADOS[0], "Canal_Tipo": CANAL_TIPO[0]}

    def extra_buttons(self, bf):
        ttk.Button(bf, text="⏱ 1ª respuesta ahora", command=lambda: self.set("Hora_Primera_Respuesta", now_s())).pack(side="left", padx=2)
        ttk.Button(bf, text="📨 Mensajes de derivación", command=self.show_derivation).pack(side="left", padx=2)

    def prefill(self, comercio, serie):
        self.clear()
        self.set("Comercio", comercio)
        self.set("Nro_Serie", serie)

    def validate(self, d):
        if not d["Comercio"] or not d["Nro_Serie"]:
            return "Nombre del comercio y N° de serie son obligatorios antes de registrar/diagnosticar."
        cdt = parse_dt(d["Fecha_Hora_Contacto"])
        if not cdt:
            return "La fecha y hora del contacto debe ser dd/mm/aaaa hh:mm."
        if d["Hora_Primera_Respuesta"] and not parse_resp(cdt, d["Hora_Primera_Respuesta"]):
            return "La primera respuesta debe ser hh:mm o dd/mm/aaaa hh:mm."
        if d["Estado"] not in ESTADOS:
            return "Elija un estado válido."
        if d["Estado"] == ESTADOS[1] and not d["Area_Derivacion"]:
            return "Para derivar, elija el área (Operaciones, Finanzas y Administración o Gerencia General)."
        if d["Estado"] == ESTADOS[0] and d["Area_Derivacion"]:
            return "Un caso cerrado en primer contacto no lleva área de derivación."
        return None

    def before_save(self, full, old):
        cdt = parse_dt(full["Fecha_Hora_Contacto"])
        rdt = parse_resp(cdt, full.get("Hora_Primera_Respuesta", ""))
        full["Fecha_Hora_Contacto"] = cdt.strftime(FMT)
        full["Hora_Primera_Respuesta"] = rdt.strftime(FMT) if rdt else ""
        full["Minutos_Respuesta"] = max(0, round((rdt - cdt).total_seconds() / 60)) if rdt else ""
        if full["Estado"] == ESTADOS[0] and not full.get("Fecha_Cierre"):
            full["Fecha_Cierre"] = now_s()
        if full["Estado"] == ESTADOS[1]:
            full["Fecha_Cierre"] = ""
        return full

    def after_save(self, full, is_new):
        if full["Estado"] == ESTADOS[1]:
            self.derivation_dialog(full)

    def show_derivation(self):
        d = self.collect()
        if not d["Comercio"]:
            messagebox.showinfo("Derivación", "Seleccione un caso o complete el formulario primero.")
            return
        if not d["Area_Derivacion"]:
            d["Area_Derivacion"] = "Operaciones"
        self.derivation_dialog(d)

    def derivation_dialog(self, d):
        area = d.get("Area_Derivacion") or "Operaciones"
        interno = (f"📌 Caso para {area}\n"
                   f"Comercio: {d.get('Comercio','')}\n"
                   f"N° de serie: {d.get('Nro_Serie','')}\n"
                   f"Canal: {d.get('Canal_Tipo','')}\n"
                   f"Motivo: {d.get('Motivo','')}\n"
                   f"Contacto: {d.get('Fecha_Hora_Contacto','')}\n"
                   f"Ya realizado: {d.get('Detalle_Solucion') or 'Revisión de primer nivel a distancia'}\n"
                   f"Se necesita: atención del área. Gracias 🙏")
        destino = {"Operaciones": "el equipo de Operaciones, que lo revisará",
                   "Finanzas y Administración": "el área de Finanzas y Administración, que revisará su consulta",
                   "Gerencia General": "Gerencia General, para su revisión"}.get(area, area)
        comercio = (f"Hola{' ' + d['Comercio'] if d.get('Comercio') else ''} 👋 Gracias por avisarnos. "
                    f"Ya pasamos su caso a {destino} y se comunicarán con usted. "
                    f"Si necesita algo más, escríbanos por aquí 🙌")
        top = tk.Toplevel(self)
        top.title("Mensajes de derivación")
        top.geometry("640x520")
        for titulo, txt in (("Mensaje interno para el área", interno), ("Mensaje para el comercio", comercio)):
            lf = ttk.LabelFrame(top, text=titulo)
            lf.pack(fill="both", expand=True, padx=8, pady=6)
            t = tk.Text(lf, height=8, wrap="word")
            t.insert("1.0", txt)
            t.pack(fill="both", expand=True, padx=4, pady=4)
            ttk.Button(lf, text="📋 Copiar", command=lambda t=t: self.app.copy(t.get("1.0", "end").strip())).pack(anchor="e", padx=4, pady=2)


# ----------------------------------------------------------------------------
# Bienvenidas
# ----------------------------------------------------------------------------
class BienvenidasTab(DataTab):
    sheet = "Comercios_Bienvenidas"
    fields = [
        ("Comercio", "Comercio *", "entry", None),
        ("Telefono_WhatsApp", "Teléfono / WhatsApp", "entry", None),
        ("Nro_Serie", "N° de serie *", "entry", None),
        ("Ejecutivo_Campo", "Ejecutivo de campo", "entry", None),
        ("Fecha_Instalacion", "Fecha de instalación (dd/mm/aaaa)", "entry", None),
        ("Bienvenida_Realizada", "Bienvenida realizada", "ro", ["No", "Sí"]),
        ("Observaciones_Bienvenida", "Observaciones", "text", None),
    ]
    show_cols = ["Comercio", "Telefono_WhatsApp", "Nro_Serie", "Ejecutivo_Campo", "Fecha_Instalacion", "Bienvenida_Realizada"]

    def defaults(self):
        return {"Bienvenida_Realizada": "No", "Fecha_Instalacion": datetime.now().strftime(FMT_D)}

    def row_tag(self, row):
        dt = parse_dt(row.get("Fecha_Instalacion", ""))
        if row.get("Bienvenida_Realizada") != "Sí" and dt and dt.date() < datetime.now().date():
            return ("pend",)
        return ()

    def extra_buttons(self, bf):
        ttk.Button(bf, text="📋 Mensaje de bienvenida", command=self.copy_msg).pack(side="left", padx=2)
        ttk.Label(bf, text="  (amarillo = bienvenida pendiente)").pack(side="left")

    def copy_msg(self):
        c = self.get("Comercio")
        self.app.copy(
            f"Hola{' ' + c if c else ''} 👋 Le escribimos de IAPAY. Ayer instalamos su SoundBox y queremos "
            "asegurarnos de que todo esté bien:\n"
            "1️⃣ ¿El equipo le *anuncia los pagos* por QR?\n"
            "2️⃣ ¿Usted y su personal ya saben usarlo?\n"
            "3️⃣ Guarde este número de WhatsApp: es su *soporte oficial* de IAPAY 📲\n"
            "Cualquier duda, escríbanos por aquí 🙌")

    def validate(self, d):
        if not d["Comercio"] or not d["Nro_Serie"]:
            return "Comercio y N° de serie son obligatorios."
        if d["Fecha_Instalacion"] and not parse_dt(d["Fecha_Instalacion"]):
            return "La fecha de instalación debe ser dd/mm/aaaa."
        return None


# ----------------------------------------------------------------------------
# Reactivación y filtro de prospectos
# ----------------------------------------------------------------------------
class ProspectosTab(DataTab):
    sheet = "Activacion_Prospectos"
    fields = [
        ("Comercio_Prospecto", "Comercio / prospecto *", "entry", None),
        ("Telefono", "Teléfono", "entry", None),
        ("Tipo_Gestion", "Tipo de gestión", "ro", ["Filtro prospecto", "Reactivación"]),
        ("Cobra_Con_QR", "¿Ya cobra con QR?", "ro", ["", "Sí", "No"]),
        ("Cuenta_Banco_Economico", "¿Cuenta en Banco Económico?", "ro", ["", "Sí", "No"]),
        ("Requiere_Visita_Operaciones", "Requiere visita (automático)", "ro", ["No", "Sí"]),
        ("Fecha_Agendada", "Fecha agendada por Operaciones", "entry", None),
        ("Motivo_Resultado", "Motivo / resultado de la llamada", "text", None),
    ]
    show_cols = ["Comercio_Prospecto", "Telefono", "Tipo_Gestion", "Cobra_Con_QR", "Cuenta_Banco_Economico",
                 "Requiere_Visita_Operaciones", "Fecha_Agendada"]

    def defaults(self):
        return {"Tipo_Gestion": "Filtro prospecto", "Requiere_Visita_Operaciones": "No"}

    def extra_buttons(self, bf):
        ttk.Button(bf, text="📋 Preguntas de filtro", command=lambda: self.app.copy(
            "Hola 👋 Gracias por su interés en el SoundBox de IAPAY. Dos preguntas rápidas:\n"
            "1️⃣ ¿Ya cobra con *QR* en su negocio?\n"
            "2️⃣ ¿Tiene una *cuenta en Banco Económico*?\n"
            "Con eso coordinamos los siguientes pasos 🙌")).pack(side="left", padx=2)
        ttk.Button(bf, text="📋 Guion de reactivación", command=self.copy_react).pack(side="left", padx=2)

    def copy_react(self):
        c = self.get("Comercio_Prospecto")
        self.app.copy(
            f"Hola{' ' + c if c else ''} 👋 Le escribimos de IAPAY para saber cómo le va con su SoundBox. "
            "Vimos que este mes tuvo pocos cobros con el equipo y queremos apoyarle. "
            "¿Tiene alguna duda o dificultad? Con gusto le ayudamos a revisarlo 🙌")

    def validate(self, d):
        if not d["Comercio_Prospecto"]:
            return "Indique el comercio o prospecto."
        if d["Tipo_Gestion"] == "Filtro prospecto":
            if not d["Cobra_Con_QR"] or not d["Cuenta_Banco_Economico"]:
                return "En el filtro de prospectos responda las dos preguntas: ¿cobra con QR? y ¿tiene cuenta en Banco Económico?"
            if d["Fecha_Agendada"] and not (d["Cobra_Con_QR"] == "Sí" and d["Cuenta_Banco_Economico"] == "Sí"):
                return "No agende con Operaciones hasta que el prospecto cobre con QR y tenga cuenta en Banco Económico."
        return None

    def before_save(self, full, old):
        ok = (full["Tipo_Gestion"] == "Filtro prospecto" and full["Cobra_Con_QR"] == "Sí"
              and full["Cuenta_Banco_Economico"] == "Sí")
        full["Requiere_Visita_Operaciones"] = "Sí" if ok else "No"
        return full

    def after_save(self, full, is_new):
        if full["Tipo_Gestion"] == "Filtro prospecto":
            if full["Requiere_Visita_Operaciones"] == "Sí":
                messagebox.showinfo("Filtro", "✅ Cumple el filtro. Puede coordinar la visita con Operaciones.")
            else:
                messagebox.showinfo("Filtro", "⏸ Aún no cumple el filtro (QR + cuenta en Banco Económico). No agendar afiliación.")


# ----------------------------------------------------------------------------
# Soporte Nivel 1
# ----------------------------------------------------------------------------
class SoporteTab(ttk.Frame):
    def __init__(self, app, parent):
        super().__init__(parent)
        self.app = app
        top = ttk.LabelFrame(self, text="Datos obligatorios antes de diagnosticar (SoundBox Q191 · 3.5\" · 4G Entel · 2.000 mAh)")
        top.pack(fill="x", padx=8, pady=8)
        self.com, self.ser = tk.StringVar(), tk.StringVar()
        ttk.Label(top, text="Comercio *").grid(row=0, column=0, padx=4, pady=4)
        ttk.Entry(top, textvariable=self.com, width=34).grid(row=0, column=1, padx=4)
        ttk.Label(top, text="N° de serie *").grid(row=0, column=2, padx=4)
        ttk.Entry(top, textvariable=self.ser, width=24).grid(row=0, column=3, padx=4)
        ttk.Button(top, text="➕ Crear caso con estos datos", command=self.to_case).grid(row=0, column=4, padx=8)
        self.warn = tk.Label(top, text="", fg="#B00020", font=("Segoe UI", 9, "bold"))
        self.warn.grid(row=1, column=0, columnspan=5, sticky="w", padx=4)

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=8, pady=4)
        left = ttk.LabelFrame(body, text="Respuestas rápidas (1 clic = copiar)")
        left.pack(side="left", fill="y", padx=(0, 8))
        for k in GUIDES:
            ttk.Button(left, text=k, width=30, command=lambda k=k: self.show(k)).pack(padx=6, pady=4)
        right = ttk.LabelFrame(body, text="Mensaje listo para pegar en Chatwoot / WhatsApp")
        right.pack(side="left", fill="both", expand=True)
        self.out = tk.Text(right, wrap="word", font=("Segoe UI", 11))
        self.out.pack(fill="both", expand=True, padx=4, pady=4)
        ttk.Button(right, text="📋 Copiar de nuevo", command=lambda: app.copy(self.out.get("1.0", "end").strip())).pack(anchor="e", padx=4, pady=4)
        self.check()

    def check(self):
        if not self.com.get().strip() or not self.ser.get().strip():
            self.warn.config(text="⚠️ Falta nombre del comercio o N° de serie: pídalos primero (botón 1).")
            return False
        self.warn.config(text="")
        return True

    def show(self, key):
        self.check()
        c, s = self.com.get().strip(), self.ser.get().strip()
        txt = GUIDES[key].replace("{c}", f" {c}" if c else "").replace("{s}", f" (equipo {s})" if s else "")
        self.out.delete("1.0", "end")
        self.out.insert("1.0", txt)
        self.app.copy(txt)

    def to_case(self):
        if not self.check():
            return
        self.app.casos.prefill(self.com.get().strip(), self.ser.get().strip())
        self.app.nb.select(self.app.casos)


# ----------------------------------------------------------------------------
# KPIs y cierre
# ----------------------------------------------------------------------------
class KpiTab(ttk.Frame):
    PERIODS = ["Hoy", "Últimos 7 días", "Mes actual", "Todo"]

    def __init__(self, app, parent):
        super().__init__(parent)
        self.app = app
        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="Período:").pack(side="left")
        self.per = tk.StringVar(value="Hoy")
        cb = ttk.Combobox(top, textvariable=self.per, values=self.PERIODS, state="readonly", width=16)
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        ttk.Button(top, text="🔄 Actualizar", command=self.refresh).pack(side="left")
        cards = ttk.Frame(self)
        cards.pack(fill="x", padx=8)
        self.card = {}
        for i, (k, t) in enumerate([("total", "Casos atendidos"), ("tpr", "Prom. 1ª respuesta (min)\nMeta ≤ 15"),
                                    ("fcr", "Resolución 1er contacto\nMeta ≥ 70%"), ("open", "Abiertos en Operaciones")]):
            f = tk.Frame(cards, bd=1, relief="solid", padx=10, pady=6)
            f.grid(row=0, column=i, padx=4, sticky="nsew")
            cards.columnconfigure(i, weight=1)
            lbl = tk.Label(f, text="–", font=("Segoe UI", 22, "bold"))
            lbl.pack()
            tk.Label(f, text=t, font=("Segoe UI", 9)).pack()
            self.card[k] = lbl
        self.out = tk.Text(self, height=16, wrap="word", font=("Consolas", 10))
        self.out.pack(fill="both", expand=True, padx=8, pady=8)
        bf = ttk.Frame(self)
        bf.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bf, text="📋 Copiar resumen", command=lambda: app.copy(self.out.get("1.0", "end").strip())).pack(side="left", padx=2)
        ttk.Button(bf, text="💾 Exportar a .txt", command=self.export).pack(side="left", padx=2)

    def in_period(self, dt):
        p, n = self.per.get(), datetime.now()
        if p == "Todo":
            return True
        if dt is None:
            return False
        if p == "Hoy":
            return dt.date() == n.date()
        if p.startswith("Últimos"):
            return dt.date() >= (n - timedelta(days=6)).date()
        return (dt.year, dt.month) == (n.year, n.month)

    def refresh(self):
        try:
            data = [r for _, r in read_rows("Registro_Casos")]
        except Exception as e:  # noqa
            self.app.status.set(f"No se pudo leer el Excel: {e}")
            return
        sel = [r for r in data if self.in_period(parse_dt(r["Fecha_Hora_Contacto"]))]
        total = len(sel)
        mins = []
        for r in sel:
            try:
                mins.append(float(r["Minutos_Respuesta"]))
            except ValueError:
                pass
        tpr = sum(mins) / len(mins) if mins else None
        fcr = (sum(1 for r in sel if r["Estado"] == ESTADOS[0]) * 100 / total) if total else None
        abiertos = [r for r in data if r["Estado"] == ESTADOS[1] and r["Area_Derivacion"] == "Operaciones" and not r["Fecha_Cierre"]]
        motivos = Counter(r["Motivo"] or "Sin motivo" for r in sel).most_common(5)

        ok_t = tpr is not None and tpr <= META_RESPUESTA_MIN
        ok_f = fcr is not None and fcr >= META_PRIMER_CONTACTO
        self.card["total"].config(text=str(total))
        self.card["tpr"].config(text="–" if tpr is None else f"{tpr:.1f}", fg="#1B7F3B" if ok_t else "#B00020")
        self.card["fcr"].config(text="–" if fcr is None else f"{fcr:.0f}%", fg="#1B7F3B" if ok_f else "#B00020")
        self.card["open"].config(text=str(len(abiertos)))

        L = [f"📊 CIERRE · {self.per.get()} · IAPAY Atención al Comercio ({datetime.now().strftime(FMT_D)})",
             f"• Casos atendidos: {total}",
             "• Tiempo promedio de 1ª respuesta: " + ("sin datos" if tpr is None else f"{tpr:.1f} min (meta ≤ {META_RESPUESTA_MIN}) {'✅' if ok_t else '⚠️'}"),
             "• Resolución en primer contacto: " + ("sin datos" if fcr is None else f"{fcr:.0f}% (meta ≥ {META_PRIMER_CONTACTO:.0f}%) {'✅' if ok_f else '⚠️'}"),
             "• Motivos más frecuentes: " + (", ".join(f"{m} ({n})" for m, n in motivos) or "sin datos"),
             f"• Casos abiertos para Operaciones: {len(abiertos)}"]
        L += [f"   - {r['Comercio']} (serie {r['Nro_Serie']}) – {r['Motivo'] or 'sin motivo'} – {r['Fecha_Hora_Contacto']}" for r in abiertos]
        self.out.delete("1.0", "end")
        self.out.insert("1.0", "\n".join(L))

    def export(self):
        p = filedialog.asksaveasfilename(initialdir=APP_DIR, defaultextension=".txt",
                                         initialfile=f"Cierre_{datetime.now():%Y%m%d}.txt")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(self.out.get("1.0", "end").strip())
            self.app.status.set(f"Resumen exportado: {p}")


# ----------------------------------------------------------------------------
# Ventana principal
# ----------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IAPAY · Atención al Comercio · SoundBox Q191")
        self.geometry("1120x780")
        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            pass
        self.status = tk.StringVar(value=f"Excel: {XLSX}")
        try:
            ensure_workbook()
        except Exception as e:  # noqa
            messagebox.showerror("Excel", f"No se pudo abrir/crear el Excel:\n{XLSX}\n\n{e}")
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self.soporte = SoporteTab(self, self.nb)
        self.casos = CasosTab(self, self.nb)
        self.bienv = BienvenidasTab(self, self.nb)
        self.prosp = ProspectosTab(self, self.nb)
        self.kpi = KpiTab(self, self.nb)
        for tab, name in ((self.soporte, "🛠 Soporte Nivel 1"), (self.casos, "📝 Registro y derivación"),
                          (self.bienv, "👋 Bienvenidas"), (self.prosp, "📞 Reactivación / Prospectos"),
                          (self.kpi, "📊 KPIs y cierre")):
            self.nb.add(tab, text=name)
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x", side="bottom")
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.refresh_all())
        self.refresh_all()

    def refresh_all(self):
        for t in (self.casos, self.bienv, self.prosp):
            t.refresh()
        self.kpi.refresh()

    def copy(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.status.set("📋 Copiado al portapapeles")


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()
