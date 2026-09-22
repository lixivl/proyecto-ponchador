import csv


def obtener_nombres_unicos(archivo_entrada, archivo_salida, columna='Nombre'):
    nombres_vistos = set()
    nombres_unicos = []

    with open(archivo_entrada, newline='', encoding='utf-8') as f:
        lector = csv.DictReader(f, delimiter=';')
        for fila in lector:
            nombre = fila[columna].strip()
            if nombre and nombre not in nombres_vistos:
                nombres_vistos.add(nombre)
                nombres_unicos.append(nombre)

    with open(archivo_salida, 'w', newline='', encoding='utf-8') as f:
        escritor = csv.writer(f)
        escritor.writerow([columna])
        for nombre in nombres_unicos:
            escritor.writerow([nombre])

    print(f"Se encontraron {len(nombres_unicos)} nombres únicos.")
    print(f"Guardado en: {archivo_salida}")


if __name__ == "__main__":
    obtener_nombres_unicos("datos_depurado.csv",
                           "salida.csv", columna="Nombre")
