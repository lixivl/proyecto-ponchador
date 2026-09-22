import csv


def obtener_nombres_unicos(archivo_entrada, archivo_salida, columna='Nombre'):
    registros_vistos = set()
    registros_unicos = []

    with open(archivo_entrada, newline='', encoding='utf-8') as f:
        lector = csv.DictReader(f, delimiter=';')
        for fila in lector:
            nombre = fila[columna].strip()
            if nombre and nombre not in registros_vistos:
                registros_vistos.add(nombre)
                registros_unicos.append({
                    columna: nombre,
                    'Entrada': fila.get('Entrada', '').strip(),
                    'Salida': fila.get('Salida', '').strip(),
                })

    with open(archivo_salida, 'w', newline='', encoding='utf-8') as f:
        escritor = csv.DictWriter(
            f,
            fieldnames=[columna, 'Entrada', 'Salida'],
            delimiter=';',
        )
        escritor.writeheader()
        escritor.writerows(registros_unicos)

    print(f"Se encontraron {len(registros_unicos)} nombres únicos.")
    print(f"Guardado en: {archivo_salida}")


if __name__ == "__main__":
    obtener_nombres_unicos("datos_depurado2.csv",
                           "salida.csv", columna="Nombre")
