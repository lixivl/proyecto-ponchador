import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


class Depurador:
    # Recibir en el constructor el nombre del archivo a depurar
    def __init__(self, nombre_archivo):
        self.nombre_archivo = nombre_archivo

    def depurar(self):
        # Crear df a partir del archivo CSV, tomando solo las columnas necesarias
        df_original = pd.read_csv(self.nombre_archivo, sep=";")
        df_depurado = df_original.iloc[:, [0, 2, 3]].copy()
        df_depurado.columns = ["Fecha-Hora", "id", "Nombre"]

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
        df_depurado = df_depurado[["Fecha", "Hora", "id", "Nombre"]]

        # eliminar filas duplicadas basadas en la columna "hora" "fecha" y "nombre"
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

        # Diferencia con el registro anterior del mismo día e id
        diferencia = df_depurado.groupby(["Fecha", "Nombre"])[
            "_fecha_hora"].diff()

        # El registro actual es el más reciente: eliminarlo si está dentro de 10 minutos
        filas_a_eliminar = diferencia <= pd.Timedelta(minutes=10)
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


class InterfazDepurador(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Depurador de registros de asistencia")
        self.geometry("820x600")
        self.minsize(700, 500)

        self.nombre_archivo = None
        self.df_original = None
        self.df_depurado = None

        self._construir_layout()

    # ---------------------------------------------------------------- UI --

    def _construir_layout(self):
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass

        contenedor = ttk.Frame(self, padding=16)
        contenedor.pack(fill="both", expand=True)

        # --- Barra superior: agregar archivo + depurar ---
        barra_superior = ttk.Frame(contenedor)
        barra_superior.pack(fill="x", pady=(0, 10))

        self.btn_agregar = ttk.Button(
            barra_superior, text="Agregar archivo", command=self._agregar_archivo
        )
        self.btn_agregar.pack(side="left")

        self.lbl_archivo = ttk.Label(
            barra_superior, text="Ningún archivo seleccionado", foreground="#666666"
        )
        self.lbl_archivo.pack(side="left", padx=10)

        self.btn_depurar = ttk.Button(
            barra_superior, text="Depurar", command=self._depurar, state="disabled"
        )
        self.btn_depurar.pack(side="right")

        # --- Vista previa original ---
        ttk.Label(contenedor, text="Vista previa del archivo original", font=("", 10, "bold")).pack(
            anchor="w"
        )
        self.tabla_original = self._crear_tabla(contenedor)
        self.tabla_original.pack(fill="both", expand=True, pady=(4, 14))

        # --- Barra depurado + descargar ---
        barra_depurado = ttk.Frame(contenedor)
        barra_depurado.pack(fill="x", pady=(0, 4))

        ttk.Label(
            barra_depurado, text="Vista previa del archivo depurado", font=("", 10, "bold")
        ).pack(side="left")

        self.btn_descargar = ttk.Button(
            barra_depurado, text="Descargar", command=self._descargar, state="disabled"
        )
        self.btn_descargar.pack(side="right")

        self.tabla_depurada = self._crear_tabla(contenedor)
        self.tabla_depurada.pack(fill="both", expand=True, pady=(4, 0))

        # --- Barra de estado ---
        self.lbl_estado = ttk.Label(
            self, text="Listo", anchor="w", padding=(10, 4))
        self.lbl_estado.pack(fill="x", side="bottom")

    def _crear_tabla(self, padre):
        marco = ttk.Frame(padre)
        tabla = ttk.Treeview(marco, show="headings", height=8)
        scroll_y = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
        scroll_x = ttk.Scrollbar(
            marco, orient="horizontal", command=tabla.xview)
        tabla.configure(yscrollcommand=scroll_y.set,
                        xscrollcommand=scroll_x.set)

        tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        marco.grid_rowconfigure(0, weight=1)
        marco.grid_columnconfigure(0, weight=1)

        # Guardamos el marco para poder empacarlo/mostrarlo desde fuera
        tabla.marco = marco
        return marco

    def _llenar_tabla(self, marco, df, filas_max=200):
        tabla = marco.winfo_children()[0]  # el Treeview es el primer hijo
        tabla.delete(*tabla.get_children())
        tabla["columns"] = list(df.columns)
        for col in df.columns:
            tabla.heading(col, text=col)
            tabla.column(col, width=120, anchor="w")

        for _, fila in df.head(filas_max).iterrows():
            tabla.insert("", "end", values=list(fila))

    # ---------------------------------------------------------------- Acciones --

    def _agregar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona un archivo CSV",
            filetypes=[("Archivos CSV", "*.csv"),
                       ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return

        try:
            self.df_original = pd.read_csv(ruta, sep=";")
        except Exception as e:
            messagebox.showerror("Error al leer el archivo", str(e))
            return

        self.nombre_archivo = ruta
        self.df_depurado = None
        self.lbl_archivo.config(
            text=os.path.basename(ruta), foreground="#000000")
        self._llenar_tabla(self.tabla_original, self.df_original)
        self.btn_depurar.config(state="normal")
        self.btn_descargar.config(state="disabled")
        self.tabla_depurada.winfo_children()[0].delete(
            *self.tabla_depurada.winfo_children()[0].get_children()
        )
        self.lbl_estado.config(
            text=f"Archivo cargado: {os.path.basename(ruta)}")

    def _depurar(self):
        if not self.nombre_archivo:
            return
        try:
            depurador = Depurador(self.nombre_archivo)
            self.df_depurado = depurador.depurar()
            self.df_depurado = consolidar_registros(self.df_depurado)
        except Exception as e:
            messagebox.showerror("Error al depurar", str(e))
            return

        self._llenar_tabla(self.tabla_depurada, self.df_depurado)
        self.btn_descargar.config(state="normal")
        self.lbl_estado.config(
            text=f"Depuración completa: {len(self.df_depurado)} filas resultantes"
        )

    def _descargar(self):
        if self.df_depurado is None:
            return
        nombre_sugerido = (
            os.path.splitext(os.path.basename(self.nombre_archivo))[
                0] + "_depurado.csv"
        )
        ruta_destino = filedialog.asksaveasfilename(
            title="Guardar archivo depurado",
            initialfile=nombre_sugerido,
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv")],
        )
        if not ruta_destino:
            return

        try:
            self.df_depurado.to_csv(ruta_destino, index=False, sep=";")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))
            return

        self.lbl_estado.config(text=f"Archivo guardado en: {ruta_destino}")
        messagebox.showinfo(
            "Listo", "El archivo depurado se guardó correctamente.")


def consolidar_registros(df):
    """Consolida las marcas en un registro por persona y día."""
    registros = []

    for (fecha, nombre), grupo in df.groupby(
        ["Fecha", "Nombre"], sort=False
    ):
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
                "Salida": grupo.iloc[-1]["Hora"]
                if len(grupo) > 1
                else 'Sin registro de salida',
                "id": grupo.iloc[0]["id"],
                "Nombre": nombre,
            }
        )

    consolidado = pd.DataFrame(
        registros,
        columns=["Fecha", "Entrada", "Salida", "id", "Nombre"],
    )
    consolidado["_fecha"] = pd.to_datetime(
        consolidado["Fecha"], format="%d/%m/%Y", errors="coerce"
    )
    return (
        consolidado.sort_values(
            by=["_fecha", "Entrada", "Nombre"], na_position="last"
        )
        .drop(columns=["_fecha"])
        .reset_index(drop=True)
    )


if __name__ == "__main__":
    app = InterfazDepurador()
    app.mainloop()
