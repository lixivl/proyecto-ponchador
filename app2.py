"""
Control de asistencia - Depurador, análisis y reportes quincenales de ponchadas
-------------------------------------------------------------------------------
Requisitos:
    pip install pandas pyodbc reportlab
    + Microsoft ODBC Driver 17 o 18 para SQL Server instalado en el equipo.

Tablas que usa (esquema existente):
    dbo.DatosPonchador (Fecha, Entrada, Salida, Nombre, IdEmpleado)
    dbo.Empleados      (Id, Nombre, IdHorario, IdDepartamento)
    dbo.Horarios       (IdHorario, NombreHorario, HorasTurno, EntradaOficial)
    dbo.Departamentos  (Id, Nombre)

Opcional: coloca junto a este archivo logo_izquierdo.png y logo_derecho.png
para que aparezcan en el encabezado del PDF del reporte.
"""

import calendar
import datetime as dt
import json
import math
import os
import re
import tkinter as tk
import webbrowser
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from xml.sax.saxutils import escape

import pandas as pd

try:
    import pyodbc
except ImportError:  # la app abre igual, pero avisa al conectar
    pyodbc = None


# ============================================================ Constantes ==

SIN_SALIDA = "Sin registro de salida"
SIN_HORARIO = "(Sin horario)"
SIN_DEPARTAMENTO = "(Sin departamento)"
TODOS_DEPARTAMENTOS = "(Todos los departamentos)"
ANIO_DATOS = 2025
FILAS_POR_PAGINA = 15
DIAS = ["Lunes", "Martes", "Miércoles",
        "Jueves", "Viernes", "Sábado", "Domingo"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIR_APP = Path(__file__).resolve().parent
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
        HorasTurno DECIMAL(2,1) NULL,
        EntradaOficial VARCHAR(50) NULL
    )""",
    # Si dbo.Horarios ya existía sin la columna, se agrega (no toca los datos)
    """IF OBJECT_ID('dbo.Horarios') IS NOT NULL AND COL_LENGTH('dbo.Horarios', 'EntradaOficial') IS NULL
    ALTER TABLE dbo.Horarios ADD EntradaOficial VARCHAR(50) NULL""",
    """IF OBJECT_ID('dbo.Departamentos') IS NULL
    CREATE TABLE dbo.Departamentos (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Nombre VARCHAR(200) NOT NULL
    )""",
    """IF OBJECT_ID('dbo.Empleados') IS NULL
    CREATE TABLE dbo.Empleados (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Nombre NVARCHAR(50) NULL,
        IdHorario INT NULL,
        IdDepartamento INT NULL
            CONSTRAINT FK_Empleados_Departamentos FOREIGN KEY REFERENCES dbo.Departamentos(Id)
    )""",
    # Si dbo.Empleados ya existía sin la columna, se agrega (no toca los datos)
    """IF OBJECT_ID('dbo.Empleados') IS NOT NULL AND COL_LENGTH('dbo.Empleados', 'IdDepartamento') IS NULL
    ALTER TABLE dbo.Empleados ADD IdDepartamento INT NULL""",
    """IF OBJECT_ID('dbo.DatosPonchador') IS NULL
    CREATE TABLE dbo.DatosPonchador (
        Fecha DATE NULL,
        Entrada VARCHAR(50) NULL,
        Salida VARCHAR(50) NULL,
        Nombre NVARCHAR(50) NULL,
        IdEmpleado INT NULL
    )""",
]

SQL_PONCHADAS = (
    "SELECT p.IdEmpleado, COALESCE(e.Nombre, p.Nombre) AS Empleado, e.IdHorario, "
    "       h.NombreHorario, h.HorasTurno, p.Fecha, p.Entrada, p.Salida "
    "FROM dbo.DatosPonchador p "
    "LEFT JOIN dbo.Empleados e ON e.Id = p.IdEmpleado "
    "LEFT JOIN dbo.Horarios h ON h.IdHorario = e.IdHorario "
    "WHERE p.Fecha BETWEEN CAST(? AS DATE) AND CAST(? AS DATE)"
)


# ============================================================ Configuración ==

def leer_config():
    try:
        with open(ARCHIVO_CONFIG, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def escribir_config(cfg):
    try:
        with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
    except OSError:
        pass


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
    raise ValueError(f"Hora inválida: '{txt}'. Usa el formato HH:MM")


_RE_HORA = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")


def hora_o_none(v):
    """
    Convierte lo que haya en Entrada/Salida (varchar) a time; None si no hay una hora.
    Tolera formatos como '8:05', '08:05:33', '08:05:33.0000000', '2025-01-06 14:02:00'
    y '02:05 p. m.'. El texto 'Sin registro de salida' devuelve None.
    """
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    txt = str(v).lower()
    m = _RE_HORA.search(txt)
    if not m:
        return None
    h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    resto = txt[m.end():]
    if re.search(r"\bp\.?\s*m\b", resto) and h < 12:
        h += 12
    elif re.search(r"\ba\.?\s*m\b", resto) and h == 12:
        h = 0
    if h > 23 or mi > 59 or s > 59:
        return None
    return dt.time(h, mi, s)


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


def quincenas_del_anio(anio):
    """Las 24 quincenas del año: 1-15 y 16-fin de mes."""
    lista = []
    for mes in range(1, 13):
        ultimo = calendar.monthrange(anio, mes)[1]
        lista.append((dt.date(anio, mes, 1), dt.date(anio, mes, 15)))
        lista.append((dt.date(anio, mes, 16), dt.date(anio, mes, ultimo)))
    return lista


def etiqueta_quincena(n, ini, fin):
    return f"Q{n:02d} · {ini.day} al {fin.day} de {MESES[ini.month - 1]}"


def texto_periodo(ini, fin):
    if ini.month == fin.month and ini.year == fin.year:
        return f"{ini.day} al {fin.day} de {MESES[ini.month - 1]} de {ini.year}"
    return (f"{ini.day} de {MESES[ini.month - 1]} de {ini.year} al "
            f"{fin.day} de {MESES[fin.month - 1]} de {fin.year}")


def letra_turno(nombre_horario):
    n = (nombre_horario or "").strip().upper()
    if n.startswith("MAT"):
        return "M"
    if n.startswith("VES"):
        return "V"
    return ""


# ================================================================ Análisis ==

def agrupar_dias(ponchadas):
    """Una entrada por (empleado, fecha): primera entrada y última salida válidas."""
    dias = {}
    for p in ponchadas:
        if p["Fecha"] is None:
            continue
        fecha = a_fecha(p["Fecha"])
        clave = p["IdEmpleado"] if p["IdEmpleado"] is not None else p["Empleado"]
        d = dias.setdefault((clave, fecha), dict(
            nombre=p["Empleado"], horario=p["NombreHorario"], id_horario=p["IdHorario"],
            turno=float(p["HorasTurno"]
                        ) if p["HorasTurno"] is not None else None,
            ents=[], sals=[]))
        ent, sal = hora_o_none(p["Entrada"]), hora_o_none(p["Salida"])
        if ent:
            d["ents"].append(ent)
        if sal:
            d["sals"].append(sal)
    for d in dias.values():
        ent = min(d["ents"]) if d["ents"] else None
        sal = max(d["sals"]) if d["sals"] else None
        if ent and sal and sal <= ent:
            sal = None  # una sola marca repetida: no hay salida real
        d["ent"], d["sal"] = ent, sal
    return dias


def minutos_retardo(ent, oficial, tolerancia):
    """
    Regla UTN: pasada la tolerancia se empiezan a contar los minutos de retardo.
    Ej.: oficial 08:00, tolerancia 10 -> entrada 08:24 = 14 minutos (08:10 o antes = 0).
    Se comparan horas y minutos (los segundos se ignoran).
    """
    if ent is None or oficial is None:
        return 0
    limite = oficial.hour * 60 + oficial.minute + tolerancia
    return max(0, ent.hour * 60 + ent.minute - limite)


COLS_DETALLE = ["Fecha", "Día", "Empleado", "Horario", "Entrada", "Salida", "Hrs turno",
                "Hrs trabajadas", "Diferencia (h)", "Estado entrada", "Estado turno"]
COLS_SEMANAL = ["Semana del", "Empleado", "Horario", "Días con registro", "Retardos",
                "Turnos completos", "Turnos incompletos", "Sin salida",
                "Hrs trabajadas", "Hrs esperadas", "Diferencia (h)"]
COLS_PERIODO = ["Empleado", "Horario", "Días con registro", "Retardos", "Turnos completos",
                "Turnos incompletos", "Sin salida", "Hrs trabajadas", "Hrs esperadas",
                "Diferencia (h)"]


def _fila_resumen(r):
    dif = f"{(r['trab'] - r['esperado']) / 60:+.2f}" if r["esperado"] else "—"
    return [r["dias"], r["retardos"], r["completos"], r["incompletos"], r["sin_salida"],
            fmt_horas(r["trab"]), fmt_horas(r["esperado"]), dif]


def _tag_resumen(r):
    if r["sin_salida"]:
        return "mal"
    if r["incompletos"] or r["retardos"]:
        return "alerta"
    return "neutro" if r["esperado"] == 0 else "ok"


def analizar(ponchadas, tol_turno, entradas_oficiales, tol_entrada):
    """
    Por persona y día evalúa:
      - Entrada: retardo si supera la hora oficial del horario + tolerancia
        (solo si el horario tiene hora oficial configurada).
      - Turno: horas entre entrada y salida contra HorasTurno (menos tolerancia).
    Los días sin salida no se pueden medir: se cuentan aparte y no entran en horas.
    Devuelve dict con 'detalle', 'semanal' y 'periodo' (cada uno: (filas, tags)).
    """
    dias = agrupar_dias(ponchadas)
    detalle, semanal, periodo = [], {}, {}

    def nuevo():
        return dict(nombre=None, horario=None, dias=0, retardos=0, completos=0,
                    incompletos=0, sin_salida=0, trab=0.0, esperado=0.0)

    for (clave, fecha), d in sorted(dias.items(), key=lambda kv: (kv[0][1], str(kv[1]["nombre"]))):
        ent, sal = d["ent"], d["sal"]
        trab = (a_min(sal) - a_min(ent)) if (ent and sal) else 0.0
        turno_min = d["turno"] * 60 if d["turno"] is not None else None
        oficial = entradas_oficiales.get(d["id_horario"])
        ret = minutos_retardo(ent, oficial, tol_entrada)

        if ent is None:
            est_ent = "Sin entrada"
        elif oficial is None:
            est_ent = "Sin hora oficial"
        elif ret:
            est_ent = f"Retardo ({ret} min)"
        else:
            est_ent = "A tiempo"

        lunes = fecha - dt.timedelta(days=fecha.weekday())
        accs = (semanal.setdefault((lunes, clave), nuevo()),
                periodo.setdefault(clave, nuevo()))

        def sumar(campo, valor=1):
            for r in accs:
                r[campo] += valor

        for r in accs:
            r["nombre"], r["horario"] = d["nombre"], d["horario"]
        sumar("dias")
        if ret:
            sumar("retardos")

        diferencia = "—"
        if turno_min is None:
            est_turno, tag = "Sin horario asignado", "neutro"
            if ent and sal:
                sumar("trab", trab)
        elif ent is None:
            est_turno, tag = "Entrada inválida", "mal"
        elif sal is None:
            est_turno, tag = SIN_SALIDA, "mal"
            sumar("sin_salida")
        else:
            sumar("trab", trab)
            sumar("esperado", turno_min)
            diferencia = f"{(trab - turno_min) / 60:+.2f}"
            if trab + tol_turno >= turno_min:
                est_turno, tag = "Turno completo", "ok"
                sumar("completos")
            else:
                falta = int(round(turno_min - trab))
                est_turno, tag = f"Incompleto (faltan {falta // 60} h {falta % 60} min)", "alerta"
                sumar("incompletos")
        if ret and tag in ("ok", "neutro"):
            tag = "alerta"

        detalle.append((
            [fecha.strftime("%d/%m/%Y"), DIAS[fecha.weekday()], d["nombre"], d["horario"] or "—",
             fmt_hm(ent), fmt_hm(sal) if sal else SIN_SALIDA,
             f"{d['turno']:.1f}" if d["turno"] is not None else "—",
             fmt_horas(trab), diferencia, est_ent, est_turno],
            tag))

    sem_f, sem_t = [], []
    for (lunes, _), r in sorted(semanal.items(), key=lambda kv: (kv[0][0], str(kv[1]["nombre"]))):
        sem_f.append([lunes.strftime("%d/%m/%Y"), r["nombre"],
                     r["horario"] or "—"] + _fila_resumen(r))
        sem_t.append(_tag_resumen(r))
    per_f, per_t = [], []
    for r in sorted(periodo.values(), key=lambda r: str(r["nombre"])):
        per_f.append([r["nombre"], r["horario"] or "—"] + _fila_resumen(r))
        per_t.append(_tag_resumen(r))
    return {
        "detalle": ([x[0] for x in detalle], [x[1] for x in detalle]),
        "semanal": (sem_f, sem_t),
        "periodo": (per_f, per_t),
    }


# ================================================================ Reportes ==

def construir_reporte(empleados, ponchadas, desde, hasta, entradas_oficiales,
                      tol_entrada, seleccion=None):
    """
    Arma los datos del 'Control Quincenal de Asistencia'.
    empleados: {Id: {Nombre, IdHorario, NombreHorario, Departamento}}.
    seleccion: lista de Id; si está vacía, se incluyen los empleados con ponchadas en el periodo.
    """
    dias = agrupar_dias([p for p in ponchadas if p["IdEmpleado"] is not None])
    if seleccion:
        ids = [i for i in seleccion if i in empleados]
    else:
        ids = sorted({c for (c, _) in dias if c in empleados})
    # Agrupado por departamento (cada uno inicia en su propia hoja del PDF) y luego por nombre
    ids.sort(key=lambda i: (not empleados[i].get("Departamento"),  # sin departamento al final
                            norm_nombre(empleados[i].get(
                                "Departamento") or ""),
                            norm_nombre(empleados[i]["Nombre"] or "")))

    # Columnas: lunes a viernes; sábado/domingo solo si alguien ponchó ese día
    extras = {f for (c, f) in dias if c in ids and f.weekday() >= 5}
    fechas, f = [], desde
    while f <= hasta:
        if f.weekday() < 5 or f in extras:
            fechas.append(f)
        f += dt.timedelta(days=1)

    filas = []
    for id_ in ids:
        e = empleados[id_]
        oficial = entradas_oficiales.get(e["IdHorario"])
        celdas, dias_asist, horas, retardos, omisiones = [], 0, 0.0, 0, 0
        for fecha in fechas:
            d = dias.get((id_, fecha))
            ent = d["ent"] if d else None
            sal = d["sal"] if d else None
            ret_min = minutos_retardo(ent, oficial, tol_entrada)
            tarde = ret_min > 0
            # Entrada: FIE si no hay ponchada; minutos de retardo si llegó tarde; vacía si llegó a tiempo
            # Salida: hora ponchada, o FIS si no hay ponchada
            celdas.append(dict(
                ent="FIE" if ent is None else (str(ret_min) if tarde else ""),
                sal=fmt_hm(sal) if sal else "FIS",
                tarde=tarde, ent_falta=ent is None, sal_falta=sal is None))
            if ent:
                dias_asist += 1
            else:
                omisiones += 1
            if not sal:
                omisiones += 1
            if ent and sal:
                horas += (a_min(sal) - a_min(ent)) / 60
            if tarde:
                retardos += 1
        filas.append(dict(nombre=e["Nombre"] or "", turno=letra_turno(e["NombreHorario"]),
                          departamento=e.get(
                              "Departamento") or SIN_DEPARTAMENTO,
                          celdas=celdas, dias=dias_asist, horas=horas,
                          retardos=retardos, omisiones=omisiones,
                          sin_oficial=oficial is None))
    return {"fechas": fechas, "filas": filas}


def generar_pdf(ruta, reporte, carrera, periodo_txt, tolerancia=10):
    """Genera el PDF con el formato F02PGRRH02.01 (una hoja por cada 15 profesores)."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate,
                                        Paragraph, Spacer, Table)
    except ImportError:
        raise RuntimeError(
            "Falta instalar reportlab.\n\nEjecuta:  pip install reportlab")

    W, H = landscape(letter)
    margen = 22
    filas, fechas = reporte["filas"], reporte["fechas"]
    n = len(fechas)
    # Hojas: cada departamento empieza en una hoja nueva (máx. FILAS_POR_PAGINA profesores por hoja)
    grupos = {}
    for fila in filas:
        grupos.setdefault(fila.get("departamento")
                          or SIN_DEPARTAMENTO, []).append(fila)
    hojas = [(depto, fs[i:i + FILAS_POR_PAGINA])
             for depto, fs in grupos.items() for i in range(0, len(fs), FILAS_POR_PAGINA)]
    paginas = max(1, len(hojas))

    w_nom, w_mv, w_tot = 128, 14, [22, 26, 20, 22]
    disponible = (W - 2 * margen) - w_nom - w_mv - sum(w_tot)
    w_cel = min(disponible / (2 * n), 30) if n else 30
    col_w = [w_nom, w_mv] + [w_cel] * (2 * n) + w_tot
    t0 = 2 + 2 * n

    negro = colors.black
    gris = colors.HexColor("#BFBFBF")
    gris_osc = colors.HexColor("#7F7F7F")
    ambar = colors.HexColor("#FFE08A")
    rosa = colors.HexColor("#F8CFC4")

    st_titulo = ParagraphStyle(
        "t", fontName="Helvetica-Bold", fontSize=11, leading=13, alignment=TA_CENTER)
    st_sub = ParagraphStyle("s", fontName="Helvetica-Bold",
                            fontSize=7.5, leading=9, alignment=TA_CENTER)
    st_subi = ParagraphStyle(
        "si", parent=st_sub, fontName="Helvetica-BoldOblique")
    st_carrera = ParagraphStyle(
        "c", fontName="Helvetica-BoldOblique", fontSize=11, leading=14, leftIndent=18)
    st_peq = ParagraphStyle("p", fontName="Helvetica",
                            fontSize=7, leading=9, alignment=TA_CENTER)
    st_nombre = ParagraphStyle(
        "n", fontName="Helvetica", fontSize=6.5, leading=7.2)
    st_leyenda = ParagraphStyle(
        "l", fontName="Helvetica", fontSize=6, leading=7)

    def tabla_pagina(chunk):
        enc0 = ["NOMBRE", "", "INCIDENCIAS"] + [""] * \
            (2 * n - 1) + ["TOTAL", "", "INCID.", ""]
        enc1 = ["", ""]
        for f in fechas:
            enc1 += [str(f.day), ""]
        enc1 += ["DIAS", "HORAS", "RET.", "OMIS."]
        enc2 = ["", ""] + ["ENTRADA", "SALIDA"] * n + ["", "", "", ""]
        datos, celdas_est = [enc0, enc1, enc2], []
        for i in range(FILAS_POR_PAGINA):
            r = 3 + i
            if i >= len(chunk):
                datos.append([""] * len(col_w))
                continue
            fila = chunk[i]
            row = [Paragraph(escape(fila["nombre"]), st_nombre), fila["turno"]]
            for j, c in enumerate(fila["celdas"]):
                ce, cs = 2 + 2 * j, 3 + 2 * j
                row += [c["ent"], c["sal"]]
                if c["ent_falta"]:
                    celdas_est += [("BACKGROUND", (ce, r), (ce, r), rosa),
                                   ("FONTNAME", (ce, r), (ce, r), "Helvetica-Bold")]
                elif c["tarde"]:
                    celdas_est += [("BACKGROUND", (ce, r), (ce, r), ambar),
                                   ("FONTNAME", (ce, r), (ce, r), "Helvetica-Bold")]
                if c["sal_falta"]:
                    celdas_est += [("BACKGROUND", (cs, r), (cs, r), rosa),
                                   ("FONTNAME", (cs, r), (cs, r), "Helvetica-Bold")]
            row += [str(fila["dias"]), f"{fila['horas']:.2f}",
                    str(fila["retardos"]), str(fila["omisiones"])]
            datos.append(row)

        est = [("SPAN", (0, 0), (1, 2)), ("SPAN", (2, 0), (t0 - 1, 0)),
               ("SPAN", (t0, 0), (t0 + 1, 0)), ("SPAN", (t0 + 2, 0), (t0 + 3, 0))]
        est += [("SPAN", (2 + 2 * j, 1), (3 + 2 * j, 1)) for j in range(n)]
        est += [("SPAN", (t0 + k, 1), (t0 + k, 2)) for k in range(4)]
        est += [
            ("BACKGROUND", (0, 0), (-1, 1), negro),
            ("BACKGROUND", (0, 0), (1, 2), negro),
            ("BACKGROUND", (t0, 1), (-1, 2), negro),
            ("BACKGROUND", (2, 2), (t0 - 1, 2), colors.HexColor("#EDEDED")),
            ("TEXTCOLOR", (0, 0), (-1, 1), colors.white),
            ("TEXTCOLOR", (t0, 0), (-1, 2), colors.white),
            ("TEXTCOLOR", (2, 2), (t0 - 1, 2), colors.HexColor("#444444")),
            ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7),
            ("FONTSIZE", (0, 1), (-1, 1), 6.5),
            ("FONTNAME", (2, 2), (t0 - 1, 2), "Helvetica"),
            ("FONTSIZE", (2, 2), (t0 - 1, 2), 4.5),
            ("FONTSIZE", (0, 3), (-1, -1), 6.5),
            ("LEADING", (0, 3), (-1, -1), 7),
            ("BACKGROUND", (t0, 3), (t0, -1), colors.HexColor("#D9D9D9")),
            ("GRID", (0, 0), (-1, -1), 0.4, negro),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ALIGN", (0, 3), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1),
             1), ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1),
             0.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.5),
        ] + celdas_est
        t = Table(datos, colWidths=col_w, rowHeights=[
                  11, 11, 10] + [18] * FILAS_POR_PAGINA, hAlign="CENTER")
        t.setStyle(est)
        return t

    def leyenda():
        items = [
            [("FI", gris, "Falta Injustificada"), ("FIE", gris, "Falta Injustificada Entrada"),
             ("FIS", gris, "Falta Injustificada Salida"), None],
            [("M", gris, "Matutino"), ("V", gris, "Vespertino"),
             ("EA", gris, "Entrada Autorizada"), ("SA", gris, "Salida Autorizada")],
            [("PSS", gris, "Permiso sin Goce de Sueldo"),
             ("0.00", negro, "FIE/ Horas injustificadas por omisión entrada"),
             ("0.00", gris_osc, "FIS/ Horas injustificadas por omisión salida"), None],
            [("", ambar, f"Número = minutos de retardo en la entrada, contados después de {tolerancia} min de tolerancia"),
             None, None, None],
        ]
        datos, est = [], [("FONTSIZE", (0, 0), (-1, -1), 6.5), ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                          ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN",
                                                                  (0, 0), (-1, -1), "MIDDLE"),
                          ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]
        for r, fila in enumerate(items):
            row = []
            for c, it in enumerate(fila):
                if it is None:
                    row += ["", ""]
                    continue
                sw, color, texto = it
                row += [sw, Paragraph(escape(texto), st_leyenda)]
                est.append(("BACKGROUND", (2 * c, r), (2 * c, r), color))
                if color in (negro, gris_osc):
                    est.append(("TEXTCOLOR", (2 * c, r),
                               (2 * c, r), colors.white))
            datos.append(row)
        est.append(("SPAN", (1, 3), (7, 3)))
        t = Table(datos, colWidths=[30, 150] * 4,
                  rowHeights=11, hAlign="CENTER")
        t.setStyle(est)
        return t

    def firmas():
        datos = [["Elaboró:", "", "Revisó:", "", "Recibió:"], [""] * 5,
                 ["Analista Administrativo", "", "Jefe de Departamento de Recursos Humanos", "", "Director de Carrera"]]
        t = Table(datos, colWidths=[190, 40, 190, 40, 190], rowHeights=[
                  10, 26, 10], hAlign="CENTER")
        t.setStyle([("FONTSIZE", (0, 0), (-1, -1), 6.5), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]
                   + [("LINEBELOW", (c, 1), (c, 1), 1.2, negro) for c in (0, 2, 4)])
        return t

    def logo(nombre):
        for ext in (".png", ".jpg", ".jpeg"):
            p = DIR_APP / (nombre + ext)
            if p.exists():
                return str(p)
        return None

    logo_izq, logo_der = logo("logo_izquierdo"), logo("logo_derecho")

    def pie(c, doc):
        c.saveState()
        for ruta_logo, x in ((logo_izq, margen + 30), (logo_der, W - margen - 30 - 58)):
            if ruta_logo:
                try:
                    c.drawImage(ruta_logo, x, H - 20 - 58, width=58, height=58,
                                preserveAspectRatio=True, mask="auto")
                except Exception:
                    pass
        c.setFont("Helvetica-Bold", 7)
        c.drawString(margen + 18, 30, "F02PGRRH02.01")
        caja_w, x0 = 56, W - margen - 3 * 56
        cajas = [["Fecha revisión", "02/04/2025"],
                 [f"Hoja {doc.page} de {paginas}"], ["Rev. 01"]]
        for i, lineas in enumerate(cajas):
            c.rect(x0 + i * caja_w, 18, caja_w, 22)
            c.setFont("Helvetica-Bold", 5.5)
            y = 29 + (len(lineas) - 1) * 3.5
            for linea in lineas:
                c.drawCentredString(x0 + i * caja_w + caja_w / 2, y, linea)
                y -= 7
        c.restoreState()

    story = []
    for p, (depto, chunk) in enumerate(hojas):
        story += [
            Spacer(1, 4),
            Paragraph("UNIVERSIDAD TECNOLÓGICA DE NOGALES, SONORA", st_titulo),
            Paragraph("Dirección de Administración y Finanzas", st_sub),
            Paragraph("Departamento de Recursos Humanos", st_subi),
            Paragraph(
                "Control Quincenal de Asistencia Profesor de Asignatura", st_subi),
            Spacer(1, 8),
            Paragraph(f"Carrera: {escape(carrera)}", st_carrera),
            Paragraph(f"Departamento: {escape(depto)}", st_carrera),
            Paragraph(f"Quincena del: {escape(periodo_txt)}", st_peq),
            Paragraph("PROFESORES DE ASIGNATURA", st_peq),
            Spacer(1, 4),
            tabla_pagina(chunk),
            Spacer(1, 10),
            leyenda(),
            Spacer(1, 12),
            firmas(),
        ]
        if p < paginas - 1:
            story.append(PageBreak())

    doc = BaseDocTemplate(ruta, pagesize=(W, H), title="Control Quincenal de Asistencia",
                          author="Universidad Tecnológica de Nogales")
    marco = Frame(margen, 40, W - 2 * margen, H - 40 - 20, leftPadding=0, rightPadding=0,
                  topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[marco], onPage=pie)])
    doc.build(story)


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


class SelectorPeriodo(ttk.LabelFrame):
    """Elige el periodo: una de las 24 quincenas del año o (opcional) un día / rango de fechas."""

    def __init__(self, padre, permitir_rango=True, anio=ANIO_DATOS):
        super().__init__(padre, text="Periodo", padding=12)
        self.permitir_rango = permitir_rango
        self.v_modo = tk.StringVar(value="quincena")
        self.v_anio = tk.StringVar(value=str(anio))
        self.v_desde = tk.StringVar()
        self.v_hasta = tk.StringVar()
        self._lista = []

        fila1 = ttk.Frame(self)
        fila1.pack(fill="x")
        if permitir_rango:
            ttk.Radiobutton(fila1, text="Por quincena", variable=self.v_modo, value="quincena",
                            command=self._toggle).pack(side="left")
        else:
            ttk.Label(fila1, text="Quincena",
                      style="Titulo.TLabel").pack(side="left")
        ttk.Label(fila1, text="   Año:").pack(side="left")
        self.sp_anio = ttk.Spinbox(fila1, from_=2000, to=2100, width=6,
                                   textvariable=self.v_anio, command=self._refrescar)
        self.sp_anio.pack(side="left", padx=6)
        self.sp_anio.bind("<Return>", lambda _: self._refrescar())
        self.sp_anio.bind("<FocusOut>", lambda _: self._refrescar())
        self.cb = ttk.Combobox(fila1, state="readonly", width=34)
        self.cb.pack(side="left", padx=8)
        self.cb.bind("<<ComboboxSelected>>", self._al_elegir)

        if permitir_rango:
            fila2 = ttk.Frame(self)
            fila2.pack(fill="x", pady=(10, 0))
            ttk.Radiobutton(fila2, text="Por día o rango de fechas", variable=self.v_modo,
                            value="rango", command=self._toggle).pack(side="left")
            ttk.Label(fila2, text="   Desde (DD/MM/AAAA):").pack(side="left")
            self.e_desde = ttk.Entry(
                fila2, textvariable=self.v_desde, width=12)
            self.e_desde.pack(side="left", padx=6)
            ttk.Label(fila2, text="Hasta:").pack(side="left")
            self.e_hasta = ttk.Entry(
                fila2, textvariable=self.v_hasta, width=12)
            self.e_hasta.pack(side="left", padx=6)
            ttk.Label(fila2, text="(para un solo día usa la misma fecha)",
                      style="Suave.TLabel").pack(side="left", padx=6)

        self._refrescar()
        self._toggle()

    def _refrescar(self):
        try:
            anio = int(self.v_anio.get())
        except ValueError:
            return
        self._lista = quincenas_del_anio(anio)
        self.cb["values"] = [etiqueta_quincena(
            i + 1, a, b) for i, (a, b) in enumerate(self._lista)]
        self.cb.current(max(self.cb.current(), 0))
        self._al_elegir()

    def _al_elegir(self, _=None):
        idx = self.cb.current()
        if idx < 0 or not self._lista:
            return
        ini, fin = self._lista[idx]
        self.v_desde.set(ini.strftime("%d/%m/%Y"))
        self.v_hasta.set(fin.strftime("%d/%m/%Y"))

    def _toggle(self):
        if not self.permitir_rango:
            return
        por_quincena = self.v_modo.get() == "quincena"
        self.sp_anio.config(state="normal" if por_quincena else "disabled")
        self.cb.config(state="readonly" if por_quincena else "disabled")
        self.e_desde.config(state="disabled" if por_quincena else "normal")
        self.e_hasta.config(state="disabled" if por_quincena else "normal")

    def obtener(self):
        """Devuelve (desde, hasta). Lanza ValueError con un mensaje claro si hay datos inválidos."""
        if not self.permitir_rango or self.v_modo.get() == "quincena":
            idx = self.cb.current()
            if idx < 0 or not self._lista:
                raise ValueError("Elige una quincena válida (revisa el año).")
            return self._lista[idx]
        desde, hasta = parse_fecha(
            self.v_desde.get()), parse_fecha(self.v_hasta.get())
        if hasta < desde:
            raise ValueError(
                "La fecha final no puede ser anterior a la inicial.")
        return desde, hasta


# ============================================================ Aplicación ==

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de asistencia — Depurador, análisis y reportes")
        self.geometry("1180x780")
        self.minsize(1000, 660)
        self.configure(bg=C["fondo"])

        self.db = BaseDatos()
        self.ruta_csv = None
        self.df_depurado = None
        self._empleados_cache = {}
        self._rep_ids = []
        self._exportables = {}

        self._estilos()
        self._construir_header()
        self._construir_estado()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(10, 4))
        self.tab_conexion = self._tab_conexion()
        self.tab_importar = self._tab_importar()
        self.tab_empleados = self._tab_empleados()
        self.tab_departamentos = self._tab_departamentos()
        self.tab_horarios = self._tab_horarios()
        self.tab_analisis = self._tab_analisis()
        self.tab_reportes = self._tab_reportes()
        self.nb.add(self.tab_conexion, text="  Conexión  ")
        self.nb.add(self.tab_importar, text="  Importar y depurar  ")
        self.nb.add(self.tab_empleados, text="  Empleados  ")
        self.nb.add(self.tab_departamentos, text="  Departamentos  ")
        self.nb.add(self.tab_horarios, text="  Horarios  ")
        self.nb.add(self.tab_analisis, text="  Análisis  ")
        self.nb.add(self.tab_reportes, text="  Reportes  ")
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
        for w in ("TCheckbutton", "TRadiobutton"):
            s.configure(w, background=C["fondo"])
            s.map(w, background=[("active", C["fondo"])])

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
        cfg = leer_config()
        self.v_servidor.set(cfg.get("servidor", "localhost"))
        self.v_base.set(cfg.get("base", ""))
        self.v_usuario.set(cfg.get("usuario", ""))
        self.v_windows.set(cfg.get("windows", True))
        if cfg.get("driver"):
            self.v_driver.set(cfg["driver"])
        self._toggle_auth()

    def _guardar_config(self):
        conexion = self._cfg()
        # nunca guardamos la contraseña en disco
        conexion.pop("password", None)
        cfg = leer_config()
        cfg.update(conexion)
        escribir_config(cfg)

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
                "¿Registrarlos sin horario ni departamento? (podrás asignarlos en la pestaña Empleados)",
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
        self.t_empleados = Tabla(f, ["Id", "Nombre", "Departamento", "Horario", "Horas del turno"],
                                 [60, 340, 280, 170, 110], alto=12)
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
        self.v_emp_depto = tk.StringVar(value=SIN_DEPARTAMENTO)
        ttk.Label(form, text="Departamento:").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        self.cb_emp_depto = ttk.Combobox(form, textvariable=self.v_emp_depto,
                                         values=[SIN_DEPARTAMENTO], state="readonly", width=60)
        self.cb_emp_depto.grid(row=1, column=1, columnspan=3,
                               sticky="w", padx=8, pady=(8, 0))

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

    def _cargar_combo_departamentos(self):
        filas = self.db.consulta(
            "SELECT Id, Nombre FROM dbo.Departamentos ORDER BY Nombre")
        valores = [SIN_DEPARTAMENTO] + \
            [f"{r['Id']} · {r['Nombre']}" for r in filas]
        self.cb_emp_depto["values"] = valores
        if self.v_emp_depto.get() not in valores:
            self.v_emp_depto.set(SIN_DEPARTAMENTO)

    def _cargar_empleados(self):
        if not self.db.conectada:
            return
        try:
            self._cargar_combo_horarios()
            self._cargar_combo_departamentos()
            filas = self.db.consulta(
                "SELECT e.Id, e.Nombre, e.IdHorario, h.NombreHorario, h.HorasTurno, "
                "       e.IdDepartamento, d.Nombre AS Departamento "
                "FROM dbo.Empleados e "
                "LEFT JOIN dbo.Horarios h ON h.IdHorario = e.IdHorario "
                "LEFT JOIN dbo.Departamentos d ON d.Id = e.IdDepartamento "
                "ORDER BY e.Nombre")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron leer los empleados:\n\n{e}")
            return
        self._empleados_cache = {r["Id"]: r for r in filas}
        self.t_empleados.cargar(
            [[r["Id"], r["Nombre"], r["Departamento"] or "—", r["NombreHorario"] or "—",
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
        elegido = SIN_DEPARTAMENTO
        for v in self.cb_emp_depto["values"]:
            if r["IdDepartamento"] is not None and v.startswith(f"{r['IdDepartamento']} · "):
                elegido = v
                break
        self.v_emp_depto.set(elegido)

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
        return (nombre, self._id_de_combo(self.v_emp_horario.get()),
                self._id_de_combo(self.v_emp_depto.get()))

    def _emp_agregar(self):
        if not self._requiere_conexion():
            return
        datos = self._datos_empleado()
        if not datos:
            return
        try:
            self.db.ejecutar(
                "INSERT INTO dbo.Empleados (Nombre, IdHorario, IdDepartamento) VALUES (?, ?, ?)", datos)
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
            self.db.ejecutar(
                "UPDATE dbo.Empleados SET Nombre = ?, IdHorario = ?, IdDepartamento = ? WHERE Id = ?",
                (datos[0], datos[1], datos[2], int(sel[0])))
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

    # ------------------------------------------------- Pestaña: Departamentos --

    def _tab_departamentos(self):
        f = ttk.Frame(self.nb, padding=14)
        ttk.Label(f, text="Catálogo de departamentos. Luego se asigna a cada empleado en la pestaña "
                          "Empleados y sirve para filtrar el reporte.",
                  style="Suave.TLabel").pack(anchor="w", pady=(0, 8))

        self.t_departamentos = Tabla(
            f, ["Id", "Nombre", "Empleados"], [70, 640, 110], alto=9)
        self.t_departamentos.pack(fill="both", expand=True)
        self.t_departamentos.tree.bind(
            "<<TreeviewSelect>>", self._departamento_seleccionado)

        form = ttk.LabelFrame(f, text="Datos del departamento", padding=12)
        form.pack(fill="x", pady=(10, 0))
        self.v_dep_nombre = tk.StringVar()
        ttk.Label(form, text="Nombre (máx. 200):").grid(
            row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.v_dep_nombre,
                  width=70).grid(row=0, column=1, padx=8)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=10)
        boton(barra, "Agregar", self._dep_agregar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Actualizar seleccionado",
              self._dep_actualizar).pack(side="left", padx=8)
        boton(barra, "Eliminar", self._dep_eliminar,
              "Peligro.TButton").pack(side="left")
        boton(barra, "Nuevo", self._dep_nuevo).pack(side="right")
        return f

    def _cargar_departamentos(self):
        if not self.db.conectada:
            return
        try:
            filas = self.db.consulta(
                "SELECT d.Id, d.Nombre, COUNT(e.Id) AS Empleados "
                "FROM dbo.Departamentos d LEFT JOIN dbo.Empleados e ON e.IdDepartamento = d.Id "
                "GROUP BY d.Id, d.Nombre ORDER BY d.Nombre")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron leer los departamentos:\n\n{e}")
            return
        self.t_departamentos.cargar([[r["Id"], r["Nombre"], r["Empleados"]] for r in filas],
                                    ids=[r["Id"] for r in filas])
        self._dep_nuevo()

    def _dep_nuevo(self):
        self.t_departamentos.tree.selection_set(())
        self.v_dep_nombre.set("")

    def _departamento_seleccionado(self, _=None):
        sel = self.t_departamentos.tree.selection()
        if not sel:
            return
        self.v_dep_nombre.set(
            self.t_departamentos.tree.item(sel[0], "values")[1])

    def _datos_departamento(self, id_actual=None):
        """Valida el nombre y que no exista otro igual. Devuelve el nombre o None."""
        nombre = " ".join(self.v_dep_nombre.get().split())
        if not nombre or len(nombre) > 200:
            messagebox.showwarning("Nombre inválido",
                                   "El nombre es obligatorio y acepta máximo 200 caracteres.")
            return None
        try:
            repetido = self.db.consulta(
                "SELECT COUNT(*) AS n FROM dbo.Departamentos WHERE Nombre = ? AND Id <> ?",
                (nombre, id_actual if id_actual is not None else -1))[0]["n"]
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return None
        if repetido:
            messagebox.showwarning(
                "Departamento repetido", "Ya existe un departamento con ese nombre.")
            return None
        return nombre

    def _dep_agregar(self):
        if not self._requiere_conexion():
            return
        nombre = self._datos_departamento()
        if not nombre:
            return
        try:
            self.db.ejecutar(
                "INSERT INTO dbo.Departamentos (Nombre) VALUES (?)", (nombre,))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_departamentos()
        self._estado(f"Departamento agregado: {nombre}")

    def _dep_actualizar(self):
        sel = self.t_departamentos.tree.selection()
        if not sel:
            messagebox.showinfo("Selecciona un departamento",
                                "Elige un departamento de la lista.")
            return
        id_ = int(sel[0])
        nombre = self._datos_departamento(id_)
        if not nombre:
            return
        try:
            self.db.ejecutar(
                "UPDATE dbo.Departamentos SET Nombre = ? WHERE Id = ?", (nombre, id_))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_departamentos()
        self._estado("Departamento actualizado")

    def _dep_eliminar(self):
        sel = self.t_departamentos.tree.selection()
        if not sel:
            return
        id_ = int(sel[0])
        try:
            n = self.db.consulta("SELECT COUNT(*) AS n FROM dbo.Empleados WHERE IdDepartamento = ?",
                                 (id_,))[0]["n"]
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        if n:
            messagebox.showwarning(
                "Departamento en uso",
                f"Este departamento está asignado a {n} empleado(s). "
                "Cámbiales el departamento antes de eliminarlo.")
            return
        if not messagebox.askyesno("Eliminar departamento", "¿Eliminar el departamento seleccionado?"):
            return
        try:
            self.db.ejecutar(
                "DELETE FROM dbo.Departamentos WHERE Id = ?", (id_,))
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._cargar_departamentos()
        self._estado("Departamento eliminado")

    # ------------------------------------------------------ Pestaña: Horarios --

    def _tab_horarios(self):
        f = ttk.Frame(self.nb, padding=14)
        ttk.Label(f, text="Cada horario define cuántas horas debe cumplir el turno. "
                          "Luego se asigna a cada empleado en la pestaña Empleados.",
                  style="Suave.TLabel").pack(anchor="w", pady=(0, 8))

        self.t_horarios = Tabla(f, ["Id", "Horario", "Horas del turno", "Entrada oficial"], [
                                70, 260, 130, 140], alto=9)
        self.t_horarios.pack(fill="both", expand=True)
        self.t_horarios.tree.bind(
            "<<TreeviewSelect>>", self._horario_seleccionado)

        form = ttk.LabelFrame(f, text="Datos del horario", padding=12)
        form.pack(fill="x", pady=(10, 0))
        self.v_hor_id = tk.StringVar()
        self.v_hor_nombre = tk.StringVar()
        self.v_hor_horas = tk.StringVar(value="8.0")
        self.v_hor_ent = tk.StringVar()
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
        ttk.Label(form, text="Entrada oficial (HH:MM):").grid(
            row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(form, textvariable=self.v_hor_ent, width=10).grid(
            row=1, column=3, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(form, style="Suave.TLabel",
                  text="Opcional. Se guarda en dbo.Horarios (columna EntradaOficial) y sirve para calcular los retardos."
                  ).grid(row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))

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

    def _entradas_oficiales(self):
        """{IdHorario: time} leído de dbo.Horarios.EntradaOficial (varchar). Vacío/ilegible = sin hora oficial."""
        res = {}
        for r in self.db.consulta("SELECT IdHorario, EntradaOficial FROM dbo.Horarios"):
            hora = hora_o_none(r["EntradaOficial"])
            if hora:
                res[r["IdHorario"]] = hora
        return res

    def _cargar_horarios(self):
        if not self.db.conectada:
            return
        try:
            filas = self.db.consulta(
                "SELECT IdHorario, NombreHorario, HorasTurno, EntradaOficial "
                "FROM dbo.Horarios ORDER BY IdHorario")
        except Exception as e:
            messagebox.showerror(
                "Error", f"No se pudieron leer los horarios:\n\n{e}")
            return

        def texto_entrada(v):
            hora = hora_o_none(v)
            if hora:
                return fmt_hm(hora)
            return str(v).strip() if v is not None and str(v).strip() else "—"

        self.t_horarios.cargar(
            [[r["IdHorario"], r["NombreHorario"], r["HorasTurno"], texto_entrada(r["EntradaOficial"])]
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
        self.v_hor_ent.set("")

    def _horario_seleccionado(self, _=None):
        sel = self.t_horarios.tree.selection()
        if not sel:
            return
        vals = self.t_horarios.tree.item(sel[0], "values")
        self.v_hor_id.set(vals[0])
        self.v_hor_nombre.set(vals[1])
        self.v_hor_horas.set(vals[2])
        self.v_hor_ent.set("" if vals[3] == "—" else vals[3])

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
            ent_txt = self.v_hor_ent.get().strip()
            entrada = parse_hora(ent_txt).strftime(
                "%H:%M") if ent_txt else None
        except ValueError as e:
            messagebox.showwarning("Dato inválido", str(e))
            return None
        return id_, nombre, horas, entrada

    def _hor_agregar(self):
        if not self._requiere_conexion():
            return
        datos = self._datos_horario(con_id=True)
        if not datos:
            return
        id_, nombre, horas, entrada = datos
        try:
            self.db.ejecutar(
                "INSERT INTO dbo.Horarios (IdHorario, NombreHorario, HorasTurno, EntradaOficial) "
                "VALUES (?, ?, ?, ?)",
                (id_, nombre, horas, entrada))
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
        _, nombre, horas, entrada = datos
        try:
            self.db.ejecutar(
                "UPDATE dbo.Horarios SET NombreHorario = ?, HorasTurno = ?, EntradaOficial = ? "
                "WHERE IdHorario = ?",
                (nombre, horas, entrada, int(sel[0])))
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

        self.sel_analisis = SelectorPeriodo(f, permitir_rango=True)
        self.sel_analisis.pack(fill="x")

        filtros = ttk.LabelFrame(f, text="Parámetros del análisis", padding=12)
        filtros.pack(fill="x", pady=(8, 0))
        self.v_tol_ent = tk.StringVar(value="10")
        self.v_tol_turno = tk.StringVar(value="10")
        self.v_an_emp = tk.StringVar(value="Todos")

        ttk.Label(filtros, text="Empleado:").grid(row=0, column=0, sticky="w")
        self.cb_an_emp = ttk.Combobox(
            filtros, textvariable=self.v_an_emp, state="readonly", width=34)
        self.cb_an_emp.grid(row=0, column=1, padx=8)
        ttk.Label(filtros, text="Tolerancia de retardo (min):").grid(
            row=0, column=2, sticky="w", padx=(16, 0))
        ttk.Spinbox(filtros, from_=0, to=60, width=5,
                    textvariable=self.v_tol_ent).grid(row=0, column=3, padx=8)
        ttk.Label(filtros, text="Tolerancia de turno completo (min):").grid(
            row=0, column=4, sticky="w", padx=(16, 0))
        ttk.Spinbox(filtros, from_=0, to=120, width=5,
                    textvariable=self.v_tol_turno).grid(row=0, column=5, padx=8)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=8)
        boton(barra, "Analizar", self._analizar,
              "Accent.TButton").pack(side="left")
        boton(barra, "Exportar pestaña actual a CSV",
              self._exportar_analisis).pack(side="left", padx=8)
        for texto, color in (("Todo en orden", C["ok"]), ("Retardo / turno incompleto", C["alerta"]),
                             ("Sin salida / error", C["mal"]), ("Sin horario", C["neutro"])):
            tk.Label(barra, text=f"  {texto}  ", bg=color, fg=C["texto"],
                     font=("Segoe UI", 9)).pack(side="right", padx=3)

        self.nb_res = ttk.Notebook(f)
        self.nb_res.pack(fill="both", expand=True)
        self.t_periodo = Tabla(self.nb_res, COLS_PERIODO,
                               [250, 110, 120, 80, 120, 130, 90, 110, 110, 110], alto=9)
        self.t_semanal = Tabla(self.nb_res, COLS_SEMANAL,
                               [100, 250, 110, 120, 80, 120, 130, 90, 110, 110, 110], alto=9)
        self.t_detalle = Tabla(self.nb_res, COLS_DETALLE,
                               [90, 90, 250, 110, 70, 150, 80, 110, 110, 150, 240], alto=9)
        self.nb_res.add(self.t_periodo, text="  Resumen del periodo  ")
        self.nb_res.add(self.t_semanal, text="  Resumen semanal  ")
        self.nb_res.add(self.t_detalle, text="  Detalle diario  ")
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

    def _consultar_ponchadas(self, desde, hasta, id_emp=None):
        sql = SQL_PONCHADAS + \
            (" AND p.IdEmpleado = ?" if id_emp is not None else "")
        params = [desde.isoformat(), hasta.isoformat()] + \
            ([id_emp] if id_emp is not None else [])
        return self.db.consulta(sql, params)

    def _analizar(self):
        if not self._requiere_conexion():
            return
        try:
            desde, hasta = self.sel_analisis.obtener()
        except ValueError as e:
            messagebox.showwarning("Periodo inválido", str(e))
            return
        try:
            tol_ent, tol_turno = int(self.v_tol_ent.get()), int(
                self.v_tol_turno.get())
        except ValueError:
            messagebox.showwarning(
                "Dato inválido", "Las tolerancias deben ser números enteros de minutos.")
            return

        try:
            self._estado("Analizando…")
            ponchadas = self._consultar_ponchadas(
                desde, hasta, self._id_de_combo(self.v_an_emp.get()))
            res = analizar(ponchadas, tol_turno,
                           self._entradas_oficiales(), tol_ent)
        except Exception as e:
            self._estado("Error en el análisis")
            messagebox.showerror("Error en el análisis", str(e))
            return

        self.t_periodo.cargar(*res["periodo"])
        self.t_semanal.cargar(*res["semanal"])
        self.t_detalle.cargar(*res["detalle"])
        self._exportables = {0: (COLS_PERIODO, res["periodo"][0], "resumen_periodo.csv"),
                             1: (COLS_SEMANAL, res["semanal"][0], "resumen_semanal.csv"),
                             2: (COLS_DETALLE, res["detalle"][0], "detalle_diario.csv")}
        self._estado(f"Análisis del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}: "
                     f"{len(res['detalle'][0])} registros diarios, {len(res['periodo'][0])} empleados")

    def _exportar_analisis(self):
        cols, filas, nombre = self._exportables.get(
            self.nb_res.index(self.nb_res.select()), (None, None, None))
        if not filas:
            messagebox.showinfo("Sin datos", "Primero ejecuta el análisis.")
            return
        destino = filedialog.asksaveasfilename(
            title="Exportar análisis", initialfile=nombre,
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

    # ------------------------------------------------------ Pestaña: Reportes --

    def _tab_reportes(self):
        f = ttk.Frame(self.nb, padding=14)

        self.sel_reporte = SelectorPeriodo(f, permitir_rango=False)
        self.sel_reporte.pack(fill="x")

        medio = ttk.Frame(f)
        medio.pack(fill="x", pady=8)

        datos = ttk.LabelFrame(medio, text="Datos del reporte", padding=12)
        datos.pack(side="left", fill="y")
        self.v_carrera = tk.StringVar()
        self.v_rep_tol = tk.StringVar(value="10")
        ttk.Label(datos, text="Carrera:").grid(row=0, column=0, sticky="w")
        ttk.Entry(datos, textvariable=self.v_carrera, width=34).grid(
            row=0, column=1, padx=8, pady=3)
        ttk.Label(datos, text="Tolerancia de retardo (min):").grid(
            row=1, column=0, sticky="w")
        ttk.Spinbox(datos, from_=0, to=60, width=5, textvariable=self.v_rep_tol).grid(
            row=1, column=1, sticky="w", padx=8, pady=3)
        self.v_rep_depto = tk.StringVar(value=TODOS_DEPARTAMENTOS)
        ttk.Label(datos, text="Departamento:").grid(
            row=2, column=0, sticky="w")
        self.cb_rep_depto = ttk.Combobox(datos, textvariable=self.v_rep_depto,
                                         values=[TODOS_DEPARTAMENTOS], state="readonly", width=40)
        self.cb_rep_depto.grid(row=2, column=1, sticky="w", padx=8, pady=3)
        self.cb_rep_depto.bind("<<ComboboxSelected>>",
                               lambda _e: self._cargar_lista_profesores())
        ttk.Label(datos, style="Suave.TLabel", justify="left",
                  text="Retardo = minutos después de la tolerancia, contados\ncon la «Entrada oficial» de la pestaña Horarios.").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        prof = ttk.LabelFrame(
            medio, text="Profesores (sin selección = todos los que tengan ponchadas)", padding=12)
        prof.pack(side="left", fill="both", expand=True, padx=(10, 0))
        marco_lb = ttk.Frame(prof)
        marco_lb.pack(side="left", fill="both", expand=True)
        self.lb_prof = tk.Listbox(
            marco_lb, selectmode="extended", height=6, exportselection=False, bg="#FFFFFF",
            fg=C["texto"], selectbackground=C["primario"], selectforeground="#FFFFFF",
            highlightthickness=1, highlightbackground=C["borde"], relief="flat",
            activestyle="none", font=("Segoe UI", 10))
        sb = ttk.Scrollbar(marco_lb, orient="vertical",
                           command=self.lb_prof.yview)
        self.lb_prof.configure(yscrollcommand=sb.set)
        self.lb_prof.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        botones = ttk.Frame(prof)
        botones.pack(side="left", padx=(10, 0), anchor="n")
        boton(botones, "Seleccionar todos",
              lambda: self.lb_prof.selection_set(0, "end")).pack(fill="x")
        boton(botones, "Limpiar selección", lambda: self.lb_prof.selection_clear(
            0, "end")).pack(fill="x", pady=6)

        barra = ttk.Frame(f)
        barra.pack(fill="x", pady=(0, 8))
        boton(barra, "Vista previa", self._reporte_vista_previa).pack(side="left")
        boton(barra, "Generar PDF", self._reporte_pdf,
              "Accent.TButton").pack(side="left", padx=8)
        for texto, color in (("Todo en orden", C["ok"]), ("Con retardos", C["alerta"]), ("Con FIE / FIS", C["mal"])):
            tk.Label(barra, text=f"  {texto}  ", bg=color, fg=C["texto"],
                     font=("Segoe UI", 9)).pack(side="right", padx=3)

        self.cont_preview = ttk.Frame(f)
        self.cont_preview.pack(fill="both", expand=True)
        self.t_reporte = Tabla(self.cont_preview, ["Profesor"], alto=8)
        self.t_reporte.pack(fill="both", expand=True)
        return f

    def _cargar_filtro_departamentos(self):
        """Refresca las opciones del filtro conservando la elección si todavía existe."""
        try:
            filas = self.db.consulta(
                "SELECT Id, Nombre FROM dbo.Departamentos ORDER BY Nombre")
        except Exception:
            return
        valores = [TODOS_DEPARTAMENTOS] + \
            [f"{r['Id']} · {r['Nombre']}" for r in filas]
        self.cb_rep_depto["values"] = valores
        if self.v_rep_depto.get() not in valores:
            self.v_rep_depto.set(TODOS_DEPARTAMENTOS)

    def _cargar_lista_profesores(self):
        self._cargar_filtro_departamentos()
        id_depto = self._id_de_combo(self.v_rep_depto.get())
        try:
            if id_depto is None:
                filas = self.db.consulta(
                    "SELECT Id, Nombre FROM dbo.Empleados ORDER BY Nombre")
            else:
                filas = self.db.consulta(
                    "SELECT Id, Nombre FROM dbo.Empleados WHERE IdDepartamento = ? ORDER BY Nombre",
                    (id_depto,))
        except Exception:
            return
        self._rep_ids = [r["Id"] for r in filas]
        self.lb_prof.delete(0, "end")
        for r in filas:
            self.lb_prof.insert("end", r["Nombre"])

    def _reporte_datos(self):
        """Consulta la BD y arma el reporte. Devuelve (reporte, desde, hasta) o None."""
        if not self._requiere_conexion():
            return None
        try:
            desde, hasta = self.sel_reporte.obtener()
        except ValueError as e:
            messagebox.showwarning("Periodo inválido", str(e))
            return None
        try:
            tol = int(self.v_rep_tol.get())
        except ValueError:
            messagebox.showwarning(
                "Dato inválido", "La tolerancia debe ser un número entero de minutos.")
            return None
        try:
            self._estado("Generando reporte…")
            emps = self.db.consulta(
                "SELECT e.Id, e.Nombre, e.IdHorario, h.NombreHorario, "
                "       e.IdDepartamento, d.Nombre AS Departamento "
                "FROM dbo.Empleados e "
                "LEFT JOIN dbo.Horarios h ON h.IdHorario = e.IdHorario "
                "LEFT JOIN dbo.Departamentos d ON d.Id = e.IdDepartamento")
            id_depto = self._id_de_combo(self.v_rep_depto.get())
            empleados = {r["Id"]: r for r in emps
                         if id_depto is None or r["IdDepartamento"] == id_depto}
            ponchadas = self._consultar_ponchadas(desde, hasta)
            seleccion = [self._rep_ids[i] for i in self.lb_prof.curselection()]
            reporte = construir_reporte(empleados, ponchadas, desde, hasta,
                                        self._entradas_oficiales(), tol, seleccion)
        except Exception as e:
            self._estado("Error al generar el reporte")
            messagebox.showerror("Error", str(e))
            return None
        if not reporte["filas"]:
            self._estado("Sin ponchadas en el periodo")
            messagebox.showinfo(
                "Sin datos", "No hay ponchadas de los profesores elegidos en ese periodo.")
            return None
        sin = [f["nombre"] for f in reporte["filas"] if f["sin_oficial"]]
        if sin:
            vista = "\n".join(sin[:8]) + ("\n…" if len(sin) > 8 else "")
            messagebox.showwarning(
                "Retardos no calculados",
                f"{len(sin)} profesor(es) no tienen un horario con «Entrada oficial» configurada "
                f"(pestañas Empleados y Horarios), así que su entrada no puede evaluarse y saldrá vacía:\n\n{vista}")
        con_ent = sum(f["dias"] for f in reporte["filas"])
        con_sal = sum(1 for f in reporte["filas"]
                      for c in f["celdas"] if not c["sal_falta"])
        self._diag_salidas = f"{con_sal} de {con_ent} días con entrada tienen salida"
        if con_ent and not con_sal:
            messagebox.showwarning(
                "Sin salidas en la base de datos",
                f"En este periodo hay {con_ent} día(s) con entrada, pero ninguno tiene hora de salida "
                "en DatosPonchador (la columna Salida está vacía o sin una hora válida).\n\n"
                "Revisa los datos guardados para ese periodo.")
        return reporte, desde, hasta, tol

    def _reporte_vista_previa(self):
        r = self._reporte_datos()
        if not r:
            return
        reporte, desde, hasta, _ = r
        cols = ["Profesor", "Departamento", "T"]
        for f in reporte["fechas"]:
            cols += [f"{f.day:02d} Ent.", f"{f.day:02d} Sal."]
        cols += ["Días", "Horas", "Ret.", "Omis."]
        anchos = [260, 220, 30] + [78] * \
            (2 * len(reporte["fechas"])) + [55, 65, 50, 55]

        self.t_reporte.destroy()
        self.t_reporte = Tabla(self.cont_preview, cols, anchos, alto=8)
        self.t_reporte.pack(fill="both", expand=True)
        filas, tags = [], []
        for fila in reporte["filas"]:
            vals = [fila["nombre"], fila["departamento"], fila["turno"]]
            for c in fila["celdas"]:
                vals += [c["ent"], c["sal"]]
            vals += [fila["dias"], f"{fila['horas']:.2f}",
                     fila["retardos"], fila["omisiones"]]
            filas.append(vals)
            tags.append("mal" if fila["omisiones"] else (
                "alerta" if fila["retardos"] else "ok"))
        self.t_reporte.cargar(filas, tags)
        self._estado(f"Vista previa: {len(filas)} profesores, quincena del {texto_periodo(desde, hasta)}. "
                     f"{self._diag_salidas}. Entrada: minutos de retardo, vacía = a tiempo, FIE = sin entrada. Salida: FIS = sin salida")

    def _reporte_pdf(self):
        r = self._reporte_datos()
        if not r:
            return
        reporte, desde, hasta, tol = r
        destino = filedialog.asksaveasfilename(
            title="Guardar reporte quincenal",
            initialfile=f"Control_quincenal_{desde:%Y-%m-%d}_al_{hasta:%Y-%m-%d}.pdf",
            defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")])
        if not destino:
            return
        try:
            generar_pdf(destino, reporte, self.v_carrera.get(
            ).strip(), texto_periodo(desde, hasta), tol)
        except Exception as e:
            messagebox.showerror("Error al generar el PDF", str(e))
            return
        self._estado(f"PDF guardado en: {destino}")
        if messagebox.askyesno("Reporte generado", "El PDF se generó correctamente.\n¿Quieres abrirlo ahora?"):
            webbrowser.open(Path(destino).resolve().as_uri())

    # ------------------------------------------------------------- Eventos --

    def _al_cambiar_tab(self, _=None):
        if not self.db.conectada:
            return
        tab = self.nb.tab(self.nb.select(), "text").strip()
        if tab == "Empleados":
            self._cargar_empleados()
        elif tab == "Departamentos":
            self._cargar_departamentos()
        elif tab == "Horarios":
            self._cargar_horarios()
        elif tab == "Análisis":
            self._refrescar_combo_analisis()
        elif tab == "Reportes":
            self._cargar_lista_profesores()


if __name__ == "__main__":
    App().mainloop()
