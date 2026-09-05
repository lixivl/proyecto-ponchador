import os
import pandas as pd


class Depurador:
    # Recibir en el constructor el nombre del archivo a depurar
    def __init__(self, nombre_archivo):
        self.nombre_archivo = nombre_archivo

    def depurar(self):
        # Lógica para depurar el archivo
        print(f"Depurando el archivo: {self.nombre_archivo}")

        # Crear df a partir del archivo CSV, tomando solo las columnas necesarias
        df_original = pd.read_csv(self.nombre_archivo, sep=";")
        df_depurado = df_original.iloc[:, [0, 2, 3]]
        df_depurado.columns = ["Fecha-Hora", "id", "Nombre"]

        # Limpiar los datos eliminando espacios en blanco
        df_depurado["Nombre"] = df_depurado["Nombre"].str.strip()

        # Separar la columna "Fecha-Hora" en dos columnas: "Fecha" y "Hora"
        df_depurado[["Fecha", "Hora"]] = df_depurado["Fecha-Hora"].str.split(
            " ", expand=True
        )
        df_depurado = df_depurado.drop(columns=["Fecha-Hora"])
        df_depurado = df_depurado[["Fecha", "Hora", "id", "Nombre"]]

        # eliminar filas duplicadas basadas en la columna "hora" "fecha" y "nombre"
        df_depurado = df_depurado.drop_duplicates(subset=["Fecha", "Hora", "Nombre"])

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
        diferencia = df_depurado.groupby(["Fecha", "Nombre"])["_fecha_hora"].diff()

        # El registro actual es el más reciente: eliminarlo si está dentro de 10 minutos
        filas_a_eliminar = diferencia <= pd.Timedelta(minutes=10)

        df_depurado = df_depurado.loc[~filas_a_eliminar].copy()

        # Eliminar columna auxiliar
        df_depurado = df_depurado.drop(columns=["_fecha_hora"])

        # Ordenar el DataFrame por id, fecha y hora
        df_depurado = df_depurado.sort_values(by=["Fecha", "Hora", "Nombre"])

        # Guardar el DataFrame depurado en un nuevo archivo CSV
        nombre_archivo_depurado = (
            os.path.splitext(self.nombre_archivo)[0] + "_depurado.csv"
        )
        df_depurado.to_csv(nombre_archivo_depurado, index=False, sep=";")
        print(f"Archivo depurado guardado como: {nombre_archivo_depurado}")

        return df_depurado


class Procesador:
    # Recibir en el constructor el nombre del archivo a procesar
    def __init__(self, nombre_archivo):
        self.nombre_archivo = nombre_archivo


if __name__ == "__main__":
    depurador = Depurador("datos.csv")
    df = depurador.depurar()
    print(df)
