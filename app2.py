"""
Control de asistencia - Depurador y analizador de ponchadas
-----------------------------------------------------------
Requisitos:
    pip install pandas pyodbc
    + Microsoft ODBC Driver 17 o 18 para SQL Server instalado en el equipo.

Tablas que usa (esquema existente):
    dbo.DatosPonchador (Fecha, Entrada, Salida, Nombre, IdEmpleado)
    dbo.Empleados      (Id, Nombre, IdHorario)
    dbo.Horarios       (IdHorario, NombreHorario, HorasTurno)
"""

import datetime as dt
import json
import os
import tkinter as tk
from decimal import Decimal, InvalidOperation
from tkinter import ttk, filedialog, messagebox

import pandas as pd

try:
    import pyodbc
except ImportError:  # la app abre igual, pero avisa al conectar
    pyodbc = None


# ============================================================ Constantes ==

SIN_SALIDA = "Sin registro de salida"
SIN_HORARIO = "(Sin horario)"
DIAS = ["Lunes", "Martes", "Miércoles",
        "Jueves", "Viernes", "Sábado", "Domingo"]
ARCHIVO_CONFIG = os.path.join(
    os.path.expanduser("~"), ".control_asistencia.json")

C = dict(
    fondo="#F4F8FB",
    texto="#264653",
    suave="#6B8A99",
    primario="#2A9D8F",
    primario_osc="#21867A",
    acento="#E3F3F0",
    encabezado="#CFE8E4",
    borde="#D5E3EA",
    ok="#DFF3E4",
    alerta="#FFF1CC",
    mal="#FBDDD3",
    neutro="#EDF1F4",
    par="#FFFFFF",
    impar="#F3F9F8",
    peligro="#E76F51",
    peligro_osc="#CF5B3D",
)

# Solo se ejecutan si las tablas NO existen (tu base actual no se modifica)
SQL_CREAR_TABLAS = [
    """IF OBJECT_ID('dbo.Horarios') IS NULL
    CREATE TABLE dbo.Horarios (
        IdHorario INT NOT NULL PRIMARY KEY,
        NombreHorario NVARCHAR(20) NULL,
        HorasTurno DECIMAL(2,1) NULL
    )""",
    """IF OBJECT_ID('dbo.Empleados') IS NULL
    CREATE TABLE dbo.Empleados (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Nombre NVARCHAR(50) NULL,
        IdHorario INT NULL
    )""",
    """IF OBJECT_ID('dbo.DatosPonchador') IS NULL
    CREATE TABLE dbo.DatosPonchador (
        Fecha DATE NULL,
        Entrada VARCHAR(50) NULL,
        Salida VARCHAR(50) NULL,
        Nombre NVARCHAR(50) NULL,
        IdEmpleado INT NULL
    )""",
]


# ================================================== Depuración (tu lógica) ==

def leer_csv(ruta):
    """Lee el CSV probando UTF-8 y, si falla, Latin-1 (común en sistemas de reloj)."""
    try:
        return pd.read_csv(ruta, sep=";", encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(ruta, sep=";", encoding="latin-1")


class Depurador:
    # Recibir en el constructor el nombre del archivo a depurar
    def __init__(self, nombre_archivo, ventana_min=10):
        self.nombre_archivo = nombre_archivo
        self.ventana_min = ventana_min

    def depurar(self):
        # Crear df a partir del archivo CSV, tomando solo las columnas necesarias
        df_original = leer_csv(self.nombre_archivo)
        df_depurado = df_original.iloc[:, [0, 3]].copy()
        df_depurado.columns = ["Fecha-Hora", "Nombre"]

        # Normalizar nombres: quitar espacios sobrantes y convertir a mayúsculas
        df_depurado["Nombre"] = (
            df_depurado["Nombre"]
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .str.upper()
        )

        # Separar la columna "Fecha-Hora" en dos columnas: "Fecha" y "Hora"
        df_depurado[["Fecha", "Hora"]] = df_depurado["Fecha-Hora"].str.split(
            " ", expand=True
        )
        df_depurado = df_depurado.drop(columns=["Fecha-Hora"])
        df_depurado = df_depurado[["Fecha", "Hora", "Nombre"]]

        # Eliminar filas duplicadas (fecha, hora y nombre)
        df_depurado = df_depurado.drop_duplicates(
            subset=["Fecha", "Hora", "Nombre"])

        # Convertir Fecha y Hora en una fecha-hora comparable
        df_depurado["_fecha_hora"] = pd.to_datetime(
            df_depurado["Fecha"] + " " + df_depurado["Hora"],
            dayfirst=True,
            errors="coerce",
        )

        # Ordenar por grupo y de la hora más antigua a la más reciente
        df_depurado = df_depurado.sort_values(
            by=["Fecha", "Nombre", "_fecha_hora"]
        ).reset_index(drop=True)

        # Diferencia con el registro anterior del mismo día
        diferencia = df_depurado.groupby(["Fecha", "Nombre"])[
            "_fecha_hora"].diff()

        # El registro actual es el más reciente: eliminarlo si cae dentro de la ventana
        filas_a_eliminar = diferencia <= pd.Timedelta(minutes=self.ventana_min)
        df_depurado = df_depurado.loc[~filas_a_eliminar].copy()

        # Eliminar columna auxiliar
        df_depurado = df_depurado.drop(columns=["_fecha_hora"])

        # Ordenar por fecha cronológica, hora y nombre
        df_depurado = df_depurado.sort_values(
            by=["Fecha", "Hora", "Nombre"],
            key=lambda x: (
                pd.to_datetime(x, format="%d/%m/%Y", errors="coerce")
                if x.name == "Fecha"
                else x
            ),
        ).reset_index(drop=True)

        return df_depurado


def consolidar_registros(df):
    """Consolida las marcas en un registro por persona y día."""
    registros = []

    for (fecha, nombre), grupo in df.groupby(["Fecha", "Nombre"], sort=False):
        grupo = grupo.copy()
        grupo["_fecha_hora"] = pd.to_datetime(
            grupo["Fecha"] + " " + grupo["Hora"],
            dayfirst=True,
            errors="coerce",
        )
        grupo = grupo.sort_values("_fecha_hora")

        registros.append(
            {
                "Fecha": fecha,
                "Entrada": grupo.iloc[0]["Hora"],
                "Salida": grupo.iloc[-1]["Hora"] if len(grupo) > 1 else SIN_SALIDA,
                "Nombre": nombre,
            }
        )

    consolidado = pd.DataFrame(
        registros, columns=["Fecha", "Entrada", "Salida", "Nombre"]
    )
    consolidado["_fecha"] = pd.to_datetime(
        consolidado["Fecha"], format="%d/%m/%Y", errors="coerce"
    )
    return (
        consolidado.sort_values(
            by=["_fecha", "Entrada", "Nombre"], na_position="last")
        .drop(columns=["_fecha"])
        .reset_index(drop=True)
    )


# ============================================================== Utilidades ==

def norm_nombre(txt):
    return " ".join(str(txt).split()).upper()


def parse_hora(txt):
    txt = str(txt).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return dt.datetime.strptime(txt, fmt).time()
        except ValueError:
            pass
    raise ValueError(f"Hora inválida: '{txt}'")


def hora_o_none(v):
    """Convierte lo que haya en Entrada/Salida (varchar) a time; None si no es una hora."""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    try:
        return parse_hora(v)
    except ValueError:
        return None


def parse_fecha(txt):
    try:
        return dt.datetime.strptime(str(txt).strip(), "%d/%m/%Y").date()
    except ValueError:
        raise ValueError(f"Fecha inválida: '{txt}'. Usa el formato DD/MM/AAAA")


def a_fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.datetime.strptime(str(v)[:10], "%Y-%m-%d").date()


def a_min(t):
    return t.hour * 60 + t.minute + t.second / 60


def fmt_hm(t):
    return t.strftime("%H:%M") if t else "—"


def fmt_horas(minutos):
    return f"{minutos / 60:.2f}"


def parse_horas_turno(txt):
    """HorasTurno es DECIMAL(2,1): de 0.0 a 9.9 con un decimal."""
    try:
        v = Decimal(str(txt).strip().replace(
            ",", ".")).quantize(Decimal("0.1"))
    except (InvalidOperation, ValueError):
        raise ValueError(
            "Las horas del turno deben ser un número, por ejemplo 7.5")
    if not (Decimal("0") <= v <= Decimal("9.9")):
        raise ValueError(
            "Las horas del turno deben estar entre 0.0 y 9.9 (la columna es DECIMAL(2,1)).")
    return v


# ================================================================ Análisis ==

def analizar(ponchadas, tolerancia_min):
    """
    Compara, por persona y día, las horas trabajadas (salida - entrada) contra las
    HorasTurno del horario asignado al empleado.

    Devuelve (detalle_filas, detalle_tags, resumen_filas, resumen_tags).
    Los días sin salida registrada no se pueden medir: se cuentan aparte y no
    entran en las horas trabajadas ni en las esperadas.
    """
    dias = {}
    for p in ponchadas:
        if p["Fecha"] is None:
            continue
        fecha = a_fecha(p["Fecha"])
        clave = p["IdEmpleado"] if p["IdEmpleado"] is not None else p["Empleado"]
        turno = float(p["HorasTurno"]) if p["HorasTurno"] is not None else None
        d = dias.setdefault((clave, fecha), dict(
            nombre=p["Empleado"], horario=p["NombreHorario"], turno=turno, ents=[], sals=[]))
        ent, sal = hora_o_none(p["Entrada"]), hora_o_none(p["Salida"])
        if ent:
            d["ents"].append(ent)
        if sal:
            d["sals"].append(sal)

    detalle, resumen = [], {}
    for (clave, fecha), d in sorted(dias.items(), key=lambda kv: (kv[0][1], str(kv[1]["nombre"]))):
        ent = min(d["ents"]) if d["ents"] else None
        sal = max(d["sals"]) if d["sals"] else None
        if ent and sal and sal <= ent:
            sal = None  # una sola marca repetida: no hay salida real
        trab = (a_min(sal) - a_min(ent)) if (ent and sal) else 0.0
        turno_min = d["turno"] * 60 if d["turno"] is not None else None

        lunes = fecha - dt.timedelta(days=fecha.weekday())
        r = resumen.setdefault((lunes, clave), dict(
            nombre=d["nombre"], horario=d["horario"], dias=0, completos=0,
            incompletos=0, sin_salida=0, trab=0.0, esperado=0.0))
        r["dias"] += 1

        diferencia = "—"
        if turno_min is None:
            estado, tag = "Sin horario asignado", "neutro"
            if ent and sal:
                r["trab"] += trab
        elif ent is None:
            estado, tag = "Entrada inválida", "mal"
        elif sal is None:
            estado, tag = SIN_SALIDA, "mal"
            r["sin_salida"] += 1
        else:
            r["trab"] += trab
            r["esperado"] += turno_min
            diferencia = f"{(trab - turno_min) / 60:+.2f}"
            if trab + tolerancia_min >= turno_min:
                estado, tag = "Turno completo", "ok"
                r["completos"] += 1
            else:
                falta = int(round(turno_min - trab))
                estado, tag = f"Incompleto (faltan {falta // 60} h {falta % 60} min)", "alerta"
                r["incompletos"] += 1

        detalle.append((
            [fecha.strftime("%d/%m/%Y"), DIAS[fecha.weekday()], d["nombre"],
             d["horario"] or "—", fmt_hm(ent), fmt_hm(
                 sal) if sal else SIN_SALIDA,
             f"{d['turno']:.1f}" if d["turno"] is not None else "—",
             fmt_horas(trab), diferencia, estado],
            tag,
        ))

    res_filas, res_tags = [], []
    for (lunes, _), r in sorted(resumen.items(), key=lambda kv: (kv[0][0], str(kv[1]["nombre"]))):
        res_filas.append([
            lunes.strftime(
                "%d/%m/%Y"), r["nombre"], r["horario"] or "—", r["dias"],
            r["completos"], r["incompletos"], r["sin_salida"],
            fmt_horas(r["trab"]), fmt_horas(r["esperado"]),
            f"{(r['trab'] - r['esperado']) / 60:+.2f}" if r["esperado"] else "—",
        ])
        if r["sin_salida"]:
            res_tags.append("mal")
        elif r["incompletos"]:
            res_tags.append("alerta")
        elif r["esperado"] == 0:
            res_tags.append("neutro")
        else:
            res_tags.append("ok")
    return [d[0] for d in detalle], [d[1] for d in detalle], res_filas, res_tags


# ========================================================== Base de datos ==

class BaseDatos:
    def __init__(self):
        self.conn = None

    @property
    def conectada(self):
        return self.conn is not None

    def conectar(self, cfg):
        if pyodbc is None:
            raise RuntimeError(
                "Falta instalar pyodbc.\n\nEjecuta:  pip install pyodbc")
        partes = [
            f"DRIVER={{{cfg['driver']}}}",
            f"SERVER={cfg['servidor']}",
            f"DATABASE={cfg['base']}",
            "TrustServerCertificate=yes",
        ]
        if cfg["windows"]:
            partes.append("Trusted_Connection=yes")
        else:
            partes += [f"UID={cfg['usuario']}", f"PWD={cfg['password']}"]
        if self.conn:
            self.desconectar()
        self.conn = pyodbc.connect(";".join(partes), timeout=10)

    def desconectar(self):
        try:
            if self.conn:
                self.conn.close()
        finally:
            self.conn = None

    def consulta(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, fila)) for fila in cur.fetchall()]

    def ejecutar(self, sql, params=()):
        cur = self.conn.cursor()
        try:
            cur.execute(sql, params)
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise


# ============================================================ Widgets base ==

class Tabla(ttk.Frame):
    """Treeview con scrollbars y filas de color."""

    def __init__(self, padre, columnas, anchos=None, alto=8):
        super().__init__(padre)
        self.columnas = list(columnas)
        self.tree = ttk.Treeview(
            self, columns=self.columnas, show="headings", height=alto)
        sy = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        for i, col in enumerate(self.columnas):
            self.tree.heading(col, text=col)
            self.tree.column(
                col, width=anchos[i] if anchos else 110, anchor="w", stretch=True)
        for tag in ("ok", "alerta", "mal", "neutro", "par", "impar"):
            self.tree.tag_configure(tag, background=C[tag])

    def limpiar(self):
        self.tree.delete(*self.tree.get_children())

    def cargar(self, filas, etiquetas=None, maximo=None, ids=None):
        self.limpiar()
        for i, fila in enumerate(filas):
            if maximo is not None and i >= maximo:
                break
            tag = etiquetas[i] if etiquetas else (
                "par" if i % 2 == 0 else "impar")
            valores = ["" if v is None else v for v in fila]
            kw = {"iid": str(ids[i])} if ids is not None else {}
            self.tree.insert("", "end", values=valores, tags=(tag,), **kw)


def boton(padre, texto, comando, estilo="Suave.TButton", **kw):
    return ttk.Button(padre, text=texto, command=comando, style=estilo, **kw)


# ============================================================ Aplicación ==

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de asistencia — Depurador y análisis de ponchadas")
        self.geometry("1120x740")
        self.minsize(980, 640)
        self.configure(bg=C["fondo"])

        self.db = BaseDatos()
        self.ruta_csv = None
        self.df_depurado = None
        self._empleados_cache = {}
        self._det_cols = self._det_rows = self._res_cols = self._res_rows = None

        self._estilos()
        self._construir_header()
        self._construir_estado()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(10, 4))
        self.tab_conexion = self._tab_conexion()
        self.tab_importar = self._tab_importar()
        self.tab_empleados = self._tab_empleados()
        self.tab_horarios = self._tab_horarios()
        self.tab_analisis = self._tab_analisis()
        self.nb.add(self.tab_conexion, text="  Conexión  ")
        self.nb.add(self.tab_importar, text="  Importar y depurar  ")
        self.nb.add(self.tab_empleados, text="  Empleados  ")
        self.nb.add(self.tab_horarios, text="  Horarios  ")
        self.nb.add(self.tab_analisis, text="  Análisis  ")
        self.nb.select(self.tab_importar)
        self.nb.bind("<<NotebookTabChanged>>", self._al_cambiar_tab)

        self._cargar_config()

    # ---------------------------------------------------------- Estilos --

    def _estilos(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        s.configure(".", background=C["fondo"],
                    foreground=C["texto"], font=("Segoe UI", 10))
        s.configure("TFrame", background=C["fondo"])
        s.configure("TLabel", background=C["fondo"], foreground=C["texto"])
        s.configure("Titulo.TLabel", font=("Segoe UI", 11, "bold"))
        s.configure("Suave.TLabel", foreground=C["suave"])
        s.configure("TLabelframe",
                    background=C["fondo"], bordercolor=C["borde"])
        s.configure("TLabelframe.Label", background=C["fondo"],
                    foreground=C["primario_osc"], font=("Segoe UI", 10, "bold"))
        s.configure("TCheckbutton", background=C["fondo"])
        s.map("TCheckbutton", background=[("active", C["fondo"])])

        s.configure("TEntry", fieldbackground="#FFFFFF",
                    bordercolor=C["borde"], padding=4)
        s.configure("TCombobox", fieldbackground="#FFFFFF",
                    bordercolor=C["borde"], padding=4)
        s.configure("TSpinbox", fieldbackground="#FFFFFF",
                    bordercolor=C["borde"], padding=4)

        base_btn = dict(font=("Segoe UI", 10, "bold"), padding=(
            14, 7), borderwidth=0, relief="flat")
        s.configure("Accent.TButton",
                    background=C["primario"], foreground="#FFFFFF", **base_btn)
        s.map("Accent.TButton",
              background=[("active", C["primario_osc"]),
                          ("disabled", "#B9D9D5")],
              foreground=[("disabled", "#F4F8FB")])
        s.configure(
            "Suave.TButton", background=C["acento"], foreground=C["primario_osc"], **base_btn)
        s.map("Suave.TButton",
              background=[("active", C["encabezado"]),
                          ("disabled", "#EEF3F5")],
              foreground=[("disabled", "#A9BAC3")])
        s.configure("Peligro.TButton", background="#FBE3DC",
                    foreground=C["peligro_osc"], **base_btn)
        s.map("Peligro.TButton", background=[("active", "#F7CFC4")])

        s.configure(
            "TNotebook", background=C["fondo"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        s.configure("TNotebook.Tab", padding=(14, 8), background=C["acento"],
                    foreground=C["texto"], font=("Segoe UI", 10, "bold"), borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", C["primario"]),
                          ("active", C["encabezado"])],
              foreground=[("selected", "#FFFFFF")])

        s.configure("Treeview", rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF",
                    foreground=C["texto"], bordercolor=C["borde"], borderwidth=1)
        s.configure("Treeview.Heading", background=C["encabezado"], foreground=C["texto"],
                    font=("Segoe UI", 10, "bold"), padding=6, relief="flat")
        s.map("Treeview.Heading", background=[("active", "#BFE0DA")])
        s.map("Treeview", background=[("selected", C["primario"])],
              foreground=[("selected", "#FFFFFF")])

        # Arreglo conocido: los colores de 'tag' no se ven en algunas versiones de Tk
        def fijar(opcion):
            return [e for e in s.map("Treeview", query_opt=opcion)
                    if e[:2] != ("!disabled", "!selected")]
        s.map("Treeview", foreground=fijar("foreground") + [("selected", "#FFFFFF")],
              background=fijar("background") + [("selected", C["primario"])])

    def _construir_header(self):
        h = tk.Frame(self, bg=C["texto"])
        h.pack(fill="x")
        tk.Label(h, text="Control de asistencia", bg=C["texto"], fg="#FFFFFF",
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=18, pady=12)
        tk.Label(h, text="Profesores y personal administrativo", bg=C["texto"],
                 fg="#9FC3C0", font=("Segoe UI", 10)).pack(side="left")
        self.lbl_conexion = tk.Label(h, text="● Sin conexión", bg=C["texto"],
                                     fg="#F4A261", font=("Segoe UI", 10, "bold"))
        self.lbl_conexion.pack(side="right", padx=18)

    def _construir_estado(self):
        self.lbl_estado = tk.Label(self, text="Listo", anchor="w", bg=C["acento"],
                                   fg=C["texto"], padx=14, pady=5, font=("Segoe UI", 9))
        self.lbl_estado.pack(fill="x", side="bottom")

    def _estado(self, texto):
        self.lbl_estado.config(text=texto)
        self.update_idletasks()

    # ------------------------------------------------------ Pestaña: Conexión --

    def _tab_conexion(self):
        f = ttk.Frame(self.nb, padding=22)
        caja = ttk.LabelFrame(f, text="Servidor SQL Server", padding=18)
        caja.pack(anchor="nw")

        self.v_servidor = tk.StringVar(value="localhost")
        self.v_base = tk.StringVar()
        self.v_usuario = tk.StringVar()
        self.v_password = tk.StringVar()
        self.v_windows = tk.BooleanVar(value=True)
        drivers = []
        if pyodbc:
            drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
        self.v_driver = tk.StringVar(
            value=drivers[0] if drivers else "ODBC Driver 17 for SQL Server")
        for pref in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
            if pref in drivers:
                self.v_driver.set(pref)
                break

        filas = [
            ("Servidor / instancia:", ttk.Entry(caja,
             textvariable=self.v_servidor, width=38)),
            ("Base de datos:", ttk.Entry(caja, textvariable=self.v_base, width=38)),
            ("Driver ODBC:", ttk.Combobox(
                caja, textvariable=self.v_driver, values=drivers, width=36)),
        ]
        for i, (txt, w) in enumerate(filas):
            ttk.Label(caja, text=txt).grid(
                row=i, column=0, sticky="w", pady=5, padx=(0, 12))
            w.grid(row=i, column=1, sticky="w", pady=5)

        ttk.Checkbutton(caja, text="Usar autenticación de Windows", variable=self.v_windows,
                        command=self._toggle_auth).grid(row=3, column=1, sticky="w", pady=(8, 2))
        ttk.Label(caja, text="Usuario:").grid(
            row=4, column=0, sticky="w", pady=5)
        self.e_usuario = ttk.Entry(caja, textvariable=self.v_usuario, width=38)
        self.e_usuario.grid(row=4, column=1, sticky="w", pady=5)
        ttk.Label(caja, text="Contraseña:").grid(
            row=5, column=0, sticky="w", pady=5)
        self.e_password = ttk.Entry(
            caja, textvariable=self.v_password, show="•", width=38)
        self.e_password.grid(row=5, column=1, sticky="w", pady=5)
        self._toggle_auth()

        barra = ttk.Frame(f)
        barra.pack(anchor="nw", pady=16)
        boton(barra, "Conectar", self._conectar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Verificar / crear tablas",
              self._crear_tablas).pack(side="left", padx=10)
        boton(barra, "Desconectar", self._desconectar).pack(side="left")

        ttk.Label(
            f, style="Suave.TLabel", justify="left",
            text=("Tablas que usa la aplicación: DatosPonchador, Empleados y Horarios.\n"
                  "«Verificar / crear tablas» solo crea las que falten; no modifica las existentes."),
        ).pack(anchor="nw")
        return f

    def _toggle_auth(self):
        estado = "disabled" if self.v_windows.get() else "normal"
        self.e_usuario.config(state=estado)
        self.e_password.config(state=estado)

    def _cfg(self):
        return dict(servidor=self.v_servidor.get().strip(), base=self.v_base.get().strip(),
                    driver=self.v_driver.get().strip(), windows=self.v_windows.get(),
                    usuario=self.v_usuario.get().strip(), password=self.v_password.get())

    def _cargar_config(self):
        try:
            with open(ARCHIVO_CONFIG, encoding="utf-8") as fh:
                cfg = json.load(fh)
            self.v_servidor.set(cfg.get("servidor", "localhost"))
            self.v_base.set(cfg.get("base", ""))
            self.v_usuario.set(cfg.get("usuario", ""))
            self.v_windows.set(cfg.get("windows", True))
            if cfg.get("driver"):
                self.v_driver.set(cfg["driver"])
            self._toggle_auth()
        except (OSError, ValueError):
            pass

    def _guardar_config(self):
        cfg = self._cfg()
        cfg.pop("password", None)  # nunca guardamos la contraseña en disco
        try:
            with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh)
        except OSError:
            pass

    def _conectar(self):
        cfg = self._cfg()
        if not cfg["servidor"] or not cfg["base"]:
            messagebox.showwarning(
                "Faltan datos", "Indica el servidor y la base de datos.")
            return
        try:
            self._estado("Conectando…")
            self.db.conectar(cfg)
        except Exception as e:
            self._estado("No se pudo conectar")
            messagebox.showerror("Error de conexión", str(e))
            return
        self._guardar_config()
        self.lbl_conexion.config(
            text=f"● Conectado: {cfg['servidor']} / {cfg['base']}", fg="#8FE3B0")
        self._estado("Conexión establecida")
        messagebox.showinfo("Conectado", "Conexión establecida correctamente.")

    def _desconectar(self):
        self.db.desconectar()
        self.lbl_conexion.config(text="● Sin conexión", fg="#F4A261")
        self._estado("Desconectado")

    def _requiere_conexion(self):
        if not self.db.conectada:
            messagebox.showwarning(
                "Sin conexión", "Primero conéctate a SQL Server en la pestaña «Conexión».")
            self.nb.select(self.tab_conexion)
            return False
        return True

    def _crear_tablas(self):
        if not self._requiere_conexion():
            return
        if not messagebox.askyesno("Verificar tablas",
                                   "Se crearán únicamente las tablas que no existan.\n¿Continuar?"):
            return
        try:
            for sql in SQL_CREAR_TABLAS:
                self.db.ejecutar(sql)
        except Exception as e:
            messagebox.showerror("Error al crear tablas", str(e))
            return
        self._estado("Tablas verificadas")
        messagebox.showinfo("Listo", "Las tablas están listas.")

    # ------------------------------------------- Pestaña: Importar y depurar --

    def _tab_importar(self):
        f = ttk.Frame(self.nb, padding=14)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=(0, 6))
        boton(barra, "Agregar archivo CSV", self._agregar_archivo,
              "Accent.TButton").pack(side="left")
        self.lbl_archivo = ttk.Label(
            barra, text="Ningún archivo seleccionado", style="Suave.TLabel")
        self.lbl_archivo.pack(side="left", padx=12)

        param = ttk.Frame(f)
        param.pack(fill="x", pady=(0, 8))
        ttk.Label(param, text="Ignorar marcas repetidas dentro de").pack(
            side="left")
        self.v_ventana = tk.StringVar(value="10")
        ttk.Spinbox(param, from_=0, to=120, width=5,
                    textvariable=self.v_ventana).pack(side="left", padx=6)
        ttk.Label(param, text="minutos").pack(side="left")
        self.btn_depurar = boton(
            param, "Depurar", self._depurar, "Accent.TButton", state="disabled")
        self.btn_depurar.pack(side="right")

        panel = ttk.PanedWindow(f, orient="vertical")
        panel.pack(fill="both", expand=True)

        caja1 = ttk.Frame(panel)
        ttk.Label(caja1, text="Vista previa del archivo original",
                  style="Titulo.TLabel").pack(anchor="w", pady=(0, 4))
        self.t_original = Tabla(caja1, ["—"], alto=6)
        self.t_original.pack(fill="both", expand=True)
        panel.add(caja1, weight=1)

        caja2 = ttk.Frame(panel)
        cab = ttk.Frame(caja2)
        cab.pack(fill="x", pady=(8, 4))
        ttk.Label(cab, text="Vista previa del archivo depurado",
                  style="Titulo.TLabel").pack(side="left")
        self.btn_bd = boton(cab, "Guardar en base de datos",
                            self._guardar_bd, "Accent.TButton", state="disabled")
        self.btn_bd.pack(side="right")
        self.btn_csv = boton(cab, "Descargar CSV",
                             self._descargar_csv, state="disabled")
        self.btn_csv.pack(side="right", padx=8)
        self.t_depurado = Tabla(caja2, ["Fecha", "Entrada", "Salida", "Nombre"], [
                                110, 110, 170, 360], alto=6)
        self.t_depurado.pack(fill="both", expand=True)
        panel.add(caja2, weight=1)
        return f

    def _agregar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona un archivo CSV",
            filetypes=[("Archivos CSV", "*.csv"),
                       ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        try:
            df = leer_csv(ruta)
        except Exception as e:
            messagebox.showerror("Error al leer el archivo", str(e))
            return

        self.ruta_csv = ruta
        self.df_depurado = None
        self.lbl_archivo.config(
            text=os.path.basename(ruta), foreground=C["texto"])
        self._reconstruir_tabla_original(df)
        self.t_depurado.limpiar()
        self.btn_depurar.config(state="normal")
        self.btn_csv.config(state="disabled")
        self.btn_bd.config(state="disabled")
        self._estado(
            f"Archivo cargado: {os.path.basename(ruta)} ({len(df)} filas)")

    def _reconstruir_tabla_original(self, df):
        padre = self.t_original.master
        self.t_original.destroy()
        self.t_original = Tabla(padre, [str(c) for c in df.columns], alto=6)
        self.t_original.pack(fill="both", expand=True)
        self.t_original.cargar(df.head(200).fillna("").values.tolist())

    def _depurar(self):
        if not self.ruta_csv:
            return
        try:
            ventana = int(self.v_ventana.get())
        except ValueError:
            messagebox.showwarning(
                "Dato inválido", "La ventana de minutos debe ser un número entero.")
            return
        try:
            self._estado("Depurando…")
            df = Depurador(self.ruta_csv, ventana).depurar()
            self.df_depurado = consolidar_registros(df)
        except Exception as e:
            self._estado("Error al depurar")
            messagebox.showerror("Error al depurar", str(e))
            return
        self.t_depurado.cargar(self.df_depurado.head(1000).values.tolist())
        self.btn_csv.config(state="normal")
        self.btn_bd.config(state="normal")
        self._estado(
            f"Depuración completa: {len(self.df_depurado)} registros (una fila por persona y día)")

    def _descargar_csv(self):
        if self.df_depurado is None:
            return
        sugerido = os.path.splitext(os.path.basename(self.ruta_csv))[
            0] + "_depurado.csv"
        destino = filedialog.asksaveasfilename(
            title="Guardar archivo depurado", initialfile=sugerido,
            defaultextension=".csv", filetypes=[("Archivos CSV", "*.csv")],
        )
        if not destino:
            return
        try:
            self.df_depurado.to_csv(
                destino, index=False, sep=";", encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))
            return
        self._estado(f"Archivo guardado en: {destino}")
        messagebox.showinfo(
            "Listo", "El archivo depurado se guardó correctamente.")

    def _guardar_bd(self):
        """Guarda el depurado en dbo.DatosPonchador (y registra empleados nuevos en dbo.Empleados)."""
        if self.df_depurado is None or not self._requiere_conexion():
            return
        df = self.df_depurado
        nombres = sorted(df["Nombre"].dropna().unique())

        largos = [n for n in nombres if len(n) > 50]
        if largos:
            messagebox.showerror(
                "Nombres demasiado largos",
                "La columna Nombre de la base acepta máximo 50 caracteres. Revisa estos nombres:\n\n"
                + "\n".join(largos[:10]))
            return

        try:
            existentes = {norm_nombre(r["Nombre"])
                          for r in self.db.consulta("SELECT Nombre FROM dbo.Empleados") if r["Nombre"]}
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudo leer la tabla Empleados:\n\n{e}")
            return
        faltantes = [n for n in nombres if norm_nombre(n) not in existentes]
        if faltantes:
            vista = ", ".join(faltantes[:8]) + \
                ("…" if len(faltantes) > 8 else "")
            if not messagebox.askyesno(
                "Empleados nuevos",
                f"{len(faltantes)} nombre(s) no existen en la tabla Empleados:\n\n{vista}\n\n"
                "¿Registrarlos sin horario? (podrás asignarles uno en la pestaña Empleados)",
            ):
                return

        cur = self.db.conn.cursor()
        omitidas = 0
        try:
            self._estado("Guardando en la base de datos…")
            for n in faltantes:
                cur.execute("INSERT INTO dbo.Empleados (Nombre) VALUES (?)", n)
            cur.execute("SELECT Id, Nombre FROM dbo.Empleados")
            mapa = {}
            for id_, nombre in cur.fetchall():
                if nombre:
                    mapa.setdefault(norm_nombre(nombre), id_)

            for r in df.itertuples(index=False):
                try:
                    fecha = dt.datetime.strptime(
                        r.Fecha, "%d/%m/%Y").date().isoformat()
                except (ValueError, TypeError):
                    omitidas += 1
                    continue
                entrada = r.Entrada
                salida = None if r.Salida == SIN_SALIDA else r.Salida
                id_emp = mapa[norm_nombre(r.Nombre)]
                # Si ya hay registro de esa persona ese día se actualiza; si no, se inserta
                cur.execute(
                    "UPDATE dbo.DatosPonchador SET Entrada = ?, Salida = ?, Nombre = ? "
                    "WHERE IdEmpleado = ? AND Fecha = CAST(? AS DATE)",
                    entrada, salida, r.Nombre, id_emp, fecha)
                if cur.rowcount == 0:
                    cur.execute(
                        "INSERT INTO dbo.DatosPonchador (Fecha, Entrada, Salida, Nombre, IdEmpleado) "
                        "VALUES (CAST(? AS DATE), ?, ?, ?, ?)",
                        fecha, entrada, salida, r.Nombre, id_emp)
            self.db.conn.commit()
        except Exception as e:
            self.db.conn.rollback()
            self._estado("Error al guardar")
            messagebox.showerror(
                "Error al guardar en la base de datos", str(e))
            return

        guardadas = len(df) - omitidas
        extra = f"\n{omitidas} fila(s) omitidas por fecha inválida." if omitidas else ""
        self._estado(f"Guardados {guardadas} registros en DatosPonchador")
        messagebox.showinfo(
            "Listo", f"Se guardaron {guardadas} registros.{extra}")

    # ----------------------------------------------------- Pestaña: Empleados --

    def _tab_empleados(self):
        f = ttk.Frame(self.nb, padding=14)
        self.t_empleados = Tabla(f, ["Id", "Nombre", "Horario", "Horas del turno"], [
                                 60, 440, 200, 130], alto=12)
        self.t_empleados.pack(fill="both", expand=True)
        self.t_empleados.tree.bind(
            "<<TreeviewSelect>>", self._empleado_seleccionado)

        form = ttk.LabelFrame(f, text="Datos del empleado", padding=12)
        form.pack(fill="x", pady=(10, 0))
        self.v_emp_nombre = tk.StringVar()
        self.v_emp_horario = tk.StringVar(value=SIN_HORARIO)
        ttk.Label(form, text="Nombre:").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.v_emp_nombre,
                  width=44).grid(row=0, column=1, padx=8)
        ttk.Label(form, text="Horario:").grid(row=0, column=2, sticky="w")
        self.cb_emp_horario = ttk.Combobox(form, textvariable=self.v_emp_horario,
                                           values=[SIN_HORARIO], state="readonly", width=30)
        self.cb_emp_horario.grid(row=0, column=3, padx=8)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=10)
        boton(barra, "Agregar", self._emp_agregar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Actualizar", self._emp_actualizar).pack(
            side="left", padx=8)
        boton(barra, "Eliminar", self._emp_eliminar,
              "Peligro.TButton").pack(side="left")
        boton(barra, "Refrescar lista", self._cargar_empleados).pack(side="right")
        return f

    @staticmethod
    def _id_de_combo(valor):
        try:
            return int(str(valor).split(" · ")[0])
        except ValueError:
            return None

    def _cargar_combo_horarios(self):
        filas = self.db.consulta(
            "SELECT IdHorario, NombreHorario, HorasTurno FROM dbo.Horarios ORDER BY IdHorario")
        valores = [SIN_HORARIO] + [
            f"{r['IdHorario']} · {r['NombreHorario'] or 'Sin nombre'} "
            f"({r['HorasTurno'] if r['HorasTurno'] is not None else '?'} h)" for r in filas]
        self.cb_emp_horario["values"] = valores
        if self.v_emp_horario.get() not in valores:
            self.v_emp_horario.set(SIN_HORARIO)

    def _cargar_empleados(self):
        if not self.db.conectada:
            return
        try:
            self._cargar_combo_horarios()
            filas = self.db.consulta(
                "SELECT e.Id, e.Nombre, e.IdHorario, h.NombreHorario, h.HorasTurno "
                "FROM dbo.Empleados e LEFT JOIN dbo.Horarios h ON h.IdHorario = e.IdHorario "
                "ORDER BY e.Nombre")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron leer los empleados:\n\n{e}")
            return
        self._empleados_cache = {r["Id"]: r for r in filas}
        self.t_empleados.cargar(
            [[r["Id"], r["Nombre"], r["NombreHorario"] or "—",
              f"{r['HorasTurno']}" if r["HorasTurno"] is not None else "—"] for r in filas],
            ids=[r["Id"] for r in filas],
        )

    def _empleado_seleccionado(self, _=None):
        sel = self.t_empleados.tree.selection()
        if not sel:
            return
        r = self._empleados_cache.get(int(sel[0]))
        if not r:
            return
        self.v_emp_nombre.set(r["Nombre"] or "")
        elegido = SIN_HORARIO
        for v in self.cb_emp_horario["values"]:
            if r["IdHorario"] is not None and v.startswith(f"{r['IdHorario']} · "):
                elegido = v
                break
        self.v_emp_horario.set(elegido)

    def _datos_empleado(self):
        nombre = norm_nombre(self.v_emp_nombre.get())
        if not nombre:
            messagebox.showwarning(
                "Falta el nombre", "Escribe el nombre del empleado.")
            return None
        if len(nombre) > 50:
            messagebox.showwarning(
                "Nombre muy largo", "El nombre acepta máximo 50 caracteres.")
            return None
        return nombre, self._id_de_combo(self.v_emp_horario.get())

    def _emp_agregar(self):
        if not self._requiere_conexion():
            return
        datos = self._datos_empleado()
        if not datos:
            return
        try:
            self.db.ejecutar(
                "INSERT INTO dbo.Empleados (Nombre, IdHorario) VALUES (?, ?)", datos)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_empleados()
        self._estado(f"Empleado agregado: {datos[0]}")

    def _emp_actualizar(self):
        sel = self.t_empleados.tree.selection()
        if not sel:
            messagebox.showinfo("Selecciona un empleado",
                                "Elige un empleado de la lista.")
            return
        datos = self._datos_empleado()
        if not datos:
            return
        try:
            self.db.ejecutar("UPDATE dbo.Empleados SET Nombre = ?, IdHorario = ? WHERE Id = ?",
                             (datos[0], datos[1], int(sel[0])))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_empleados()
        self._estado("Empleado actualizado")

    def _emp_eliminar(self):
        sel = self.t_empleados.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(
            "Eliminar empleado",
            "Se eliminará el empleado de la tabla Empleados.\n"
            "Sus registros en DatosPonchador se conservan (con su nombre).\n¿Continuar?",
        ):
            return
        try:
            self.db.ejecutar(
                "DELETE FROM dbo.Empleados WHERE Id = ?", (int(sel[0]),))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_empleados()
        self._estado("Empleado eliminado")

    # ------------------------------------------------------ Pestaña: Horarios --

    def _tab_horarios(self):
        f = ttk.Frame(self.nb, padding=14)
        ttk.Label(f, text="Cada horario define cuántas horas debe cumplir el turno. "
                          "Luego se asigna a cada empleado en la pestaña Empleados.",
                  style="Suave.TLabel").pack(anchor="w", pady=(0, 8))

        self.t_horarios = Tabla(f, ["Id", "Horario", "Horas del turno"], [
                                80, 280, 150], alto=10)
        self.t_horarios.pack(fill="both", expand=True)
        self.t_horarios.tree.bind(
            "<<TreeviewSelect>>", self._horario_seleccionado)

        form = ttk.LabelFrame(f, text="Datos del horario", padding=12)
        form.pack(fill="x", pady=(10, 0))
        self.v_hor_id = tk.StringVar()
        self.v_hor_nombre = tk.StringVar()
        self.v_hor_horas = tk.StringVar(value="8.0")
        ttk.Label(form, text="Id:").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.v_hor_id,
                  width=8).grid(row=0, column=1, padx=8)
        ttk.Label(form, text="Nombre (máx. 20):").grid(
            row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=self.v_hor_nombre,
                  width=24).grid(row=0, column=3, padx=8)
        ttk.Label(form, text="Horas del turno (0.0 – 9.9):").grid(
            row=0, column=4, sticky="w")
        ttk.Entry(form, textvariable=self.v_hor_horas,
                  width=8).grid(row=0, column=5, padx=8)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=10)
        boton(barra, "Agregar", self._hor_agregar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Actualizar seleccionado",
              self._hor_actualizar).pack(side="left", padx=8)
        boton(barra, "Eliminar", self._hor_eliminar,
              "Peligro.TButton").pack(side="left")
        boton(barra, "Nuevo", self._hor_nuevo).pack(side="right")
        return f

    def _cargar_horarios(self):
        if not self.db.conectada:
            return
        try:
            filas = self.db.consulta(
                "SELECT IdHorario, NombreHorario, HorasTurno FROM dbo.Horarios ORDER BY IdHorario")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron leer los horarios:\n\n{e}")
            return
        self.t_horarios.cargar(
            [[r["IdHorario"], r["NombreHorario"], r["HorasTurno"]]
                for r in filas],
            ids=[r["IdHorario"] for r in filas])
        self._hor_nuevo(filas)

    def _hor_nuevo(self, filas=None):
        """Limpia el formulario y propone el siguiente Id (Horarios.IdHorario no es autonumérico)."""
        if filas is None:
            try:
                filas = self.db.consulta("SELECT IdHorario FROM dbo.Horarios")
            except Exception:
                filas = []
        siguiente = max([r["IdHorario"] for r in filas], default=0) + 1
        self.t_horarios.tree.selection_set(())
        self.v_hor_id.set(str(siguiente))
        self.v_hor_nombre.set("")
        self.v_hor_horas.set("8.0")

    def _horario_seleccionado(self, _=None):
        sel = self.t_horarios.tree.selection()
        if not sel:
            return
        vals = self.t_horarios.tree.item(sel[0], "values")
        self.v_hor_id.set(vals[0])
        self.v_hor_nombre.set(vals[1])
        self.v_hor_horas.set(vals[2])

    def _datos_horario(self, con_id):
        try:
            id_ = int(self.v_hor_id.get()) if con_id else None
        except ValueError:
            messagebox.showwarning(
                "Id inválido", "El Id del horario debe ser un número entero.")
            return None
        nombre = self.v_hor_nombre.get().strip()
        if not nombre or len(nombre) > 20:
            messagebox.showwarning(
                "Nombre inválido", "El nombre es obligatorio y acepta máximo 20 caracteres.")
            return None
        try:
            horas = parse_horas_turno(self.v_hor_horas.get())
        except ValueError as e:
            messagebox.showwarning("Horas inválidas", str(e))
            return None
        return id_, nombre, horas

    def _hor_agregar(self):
        if not self._requiere_conexion():
            return
        datos = self._datos_horario(con_id=True)
        if not datos:
            return
        try:
            self.db.ejecutar(
                "INSERT INTO dbo.Horarios (IdHorario, NombreHorario, HorasTurno) VALUES (?, ?, ?)", datos)
        except Exception as e:
            messagebox.showerror("Error (¿el Id ya existe?)", str(e))
            return
        self._cargar_horarios()
        self._estado("Horario agregado")

    def _hor_actualizar(self):
        sel = self.t_horarios.tree.selection()
        if not sel:
            messagebox.showinfo("Selecciona un horario",
                                "Elige un horario de la lista.")
            return
        datos = self._datos_horario(con_id=False)
        if not datos:
            return
        try:
            self.db.ejecutar(
                "UPDATE dbo.Horarios SET NombreHorario = ?, HorasTurno = ? WHERE IdHorario = ?",
                (datos[1], datos[2], int(sel[0])))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_horarios()
        self._estado("Horario actualizado")

    def _hor_eliminar(self):
        sel = self.t_horarios.tree.selection()
        if not sel:
            return
        id_ = int(sel[0])
        try:
            n = self.db.consulta(
                "SELECT COUNT(*) AS n FROM dbo.Empleados WHERE IdHorario = ?", (id_,))[0]["n"]
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        if n:
            messagebox.showwarning(
                "Horario en uso",
                f"Este horario está asignado a {n} empleado(s). Cámbiales el horario antes de eliminarlo.")
            return
        if not messagebox.askyesno("Eliminar horario", "¿Eliminar el horario seleccionado?"):
            return
        try:
            self.db.ejecutar(
                "DELETE FROM dbo.Horarios WHERE IdHorario = ?", (id_,))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_horarios()
        self._estado("Horario eliminado")

    # ------------------------------------------------------ Pestaña: Análisis --

    def _tab_analisis(self):
        f = ttk.Frame(self.nb, padding=14)

        filtros = ttk.LabelFrame(f, text="Parámetros del análisis", padding=12)
        filtros.pack(fill="x")
        hoy = dt.date.today()
        lunes = hoy - dt.timedelta(days=hoy.weekday())
        self.v_desde = tk.StringVar(value=lunes.strftime("%d/%m/%Y"))
        self.v_hasta = tk.StringVar(value=hoy.strftime("%d/%m/%Y"))
        self.v_tol = tk.StringVar(value="10")
        self.v_an_emp = tk.StringVar(value="Todos")

        ttk.Label(filtros, text="Desde (DD/MM/AAAA):").grid(row=0,
                                                            column=0, sticky="w")
        ttk.Entry(filtros, textvariable=self.v_desde,
                  width=12).grid(row=0, column=1, padx=8)
        ttk.Label(filtros, text="Hasta:").grid(row=0, column=2, sticky="w")
        ttk.Entry(filtros, textvariable=self.v_hasta,
                  width=12).grid(row=0, column=3, padx=8)
        ttk.Label(filtros, text="Empleado:").grid(row=0, column=4, sticky="w")
        self.cb_an_emp = ttk.Combobox(
            filtros, textvariable=self.v_an_emp, state="readonly", width=34)
        self.cb_an_emp.grid(row=0, column=5, padx=8)

        ttk.Label(filtros, text="Tolerancia (min) para dar por completo el turno:").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Spinbox(filtros, from_=0, to=120, width=5, textvariable=self.v_tol).grid(
            row=1, column=3, sticky="w", padx=8, pady=(8, 0))

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=10)
        boton(barra, "Analizar", self._analizar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Exportar pestaña actual a CSV",
              self._exportar_analisis).pack(side="left", padx=8)
        for texto, color in (("Turno completo", C["ok"]), ("Turno incompleto", C["alerta"]),
                             ("Sin salida / error", C["mal"]), ("Sin horario", C["neutro"])):
            tk.Label(barra, text=f"  {texto}  ", bg=color, fg=C["texto"],
                     font=("Segoe UI", 9)).pack(side="right", padx=3)

        self.nb_res = ttk.Notebook(f)
        self.nb_res.pack(fill="both", expand=True)
        cols_det = ["Fecha", "Día", "Empleado", "Horario", "Entrada", "Salida",
                    "Hrs turno", "Hrs trabajadas", "Diferencia (h)", "Estado"]
        self.t_detalle = Tabla(self.nb_res, cols_det, [
                               90, 90, 250, 110, 70, 150, 80, 110, 110, 240], alto=10)
        cols_res = ["Semana del", "Empleado", "Horario", "Días con registro", "Turnos completos",
                    "Turnos incompletos", "Sin salida", "Hrs trabajadas", "Hrs esperadas", "Diferencia (h)"]
        self.t_resumen = Tabla(self.nb_res, cols_res, [
                               100, 250, 110, 120, 120, 130, 90, 110, 110, 110], alto=10)
        self.nb_res.add(self.t_resumen, text="  Resumen semanal  ")
        self.nb_res.add(self.t_detalle, text="  Detalle diario  ")
        self._det_cols, self._res_cols = cols_det, cols_res
        return f

    def _refrescar_combo_analisis(self):
        try:
            filas = self.db.consulta(
                "SELECT Id, Nombre FROM dbo.Empleados ORDER BY Nombre")
        except Exception:
            return
        valores = ["Todos"] + [f"{r['Id']} · {r['Nombre']}" for r in filas]
        self.cb_an_emp["values"] = valores
        if self.v_an_emp.get() not in valores:
            self.v_an_emp.set("Todos")

    def _analizar(self):
        if not self._requiere_conexion():
            return
        try:
            desde, hasta = parse_fecha(
                self.v_desde.get()), parse_fecha(self.v_hasta.get())
        except ValueError as e:
            messagebox.showwarning("Fecha inválida", str(e))
            return
        try:
            tolerancia = int(self.v_tol.get())
        except ValueError:
            messagebox.showwarning(
                "Dato inválido", "La tolerancia debe ser un número entero de minutos.")
            return
        if hasta < desde:
            messagebox.showwarning(
                "Rango inválido", "La fecha final no puede ser anterior a la inicial.")
            return

        id_emp = self._id_de_combo(self.v_an_emp.get())
        filtro = " AND p.IdEmpleado = ?" if id_emp is not None else ""
        params = [desde.isoformat(), hasta.isoformat()] + \
            ([id_emp] if id_emp is not None else [])
        try:
            self._estado("Analizando…")
            ponchadas = self.db.consulta(
                "SELECT p.IdEmpleado, COALESCE(e.Nombre, p.Nombre) AS Empleado, "
                "       h.NombreHorario, h.HorasTurno, p.Fecha, p.Entrada, p.Salida "
                "FROM dbo.DatosPonchador p "
                "LEFT JOIN dbo.Empleados e ON e.Id = p.IdEmpleado "
                "LEFT JOIN dbo.Horarios h ON h.IdHorario = e.IdHorario "
                "WHERE p.Fecha BETWEEN CAST(? AS DATE) AND CAST(? AS DATE)" + filtro, params)
            det, det_t, res, res_t = analizar(ponchadas, tolerancia)
        except Exception as e:
            self._estado("Error en el análisis")
            messagebox.showerror("Error en el análisis", str(e))
            return

        self.t_detalle.cargar(det, det_t)
        self.t_resumen.cargar(res, res_t)
        self._det_rows, self._res_rows = det, res
        self._estado(
            f"Análisis listo: {len(det)} registros diarios, {len(res)} resúmenes semanales")

    def _exportar_analisis(self):
        resumen = self.nb_res.index(self.nb_res.select()) == 0
        cols = self._res_cols if resumen else self._det_cols
        filas = self._res_rows if resumen else self._det_rows
        if not filas:
            messagebox.showinfo("Sin datos", "Primero ejecuta el análisis.")
            return
        destino = filedialog.asksaveasfilename(
            title="Exportar análisis",
            initialfile="resumen_semanal.csv" if resumen else "detalle_diario.csv",
            defaultextension=".csv", filetypes=[("Archivos CSV", "*.csv")])
        if not destino:
            return
        try:
            pd.DataFrame(filas, columns=cols).to_csv(
                destino, index=False, sep=";", encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            return
        self._estado(f"Exportado a: {destino}")

    # ------------------------------------------------------------- Eventos --

    def _al_cambiar_tab(self, _=None):
        if not self.db.conectada:
            return
        tab = self.nb.tab(self.nb.select(), "text").strip()
        if tab == "Empleados":
            self._cargar_empleados()
        elif tab == "Horarios":
            self._cargar_horarios()
        elif tab == "Análisis":
            self._refrescar_combo_analisis()


if __name__ == "__main__":
    App().mainloop()
