import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel
file_path = "C:\\Users\\marco.avila\\Downloads\\Enero 2025.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')


# Función para extraer la última fecha de la columna F (índice 5)
def extraer_ultima_fecha(celda):
    if pd.isna(celda) or not isinstance(celda, str) or "|" not in celda:
        return None
    fechas = []
    for part in celda.split(','):
        partes = part.split('|')
        if len(partes) < 2:
            continue
        try:
            fecha = datetime.strptime(partes[1].strip(), "%m/%d/%Y %I:%M:%S %p")
            fechas.append((fecha, partes[0].strip()))
        except ValueError:
            continue
    return max(fechas, key=lambda x: x[0])[0] if fechas else None


# Aplicar la extracción en la columna F (índice 5) desde la fila 2
df.iloc[:, 7] = df.iloc[:, 5].apply(extraer_ultima_fecha)

# Convertir la columna H (índice 7) a formato fecha
df.iloc[:, 7] = pd.to_datetime(df.iloc[:, 7])
df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6])  # Asegurarse de que la columna G (índice 6) también es datetime

# Restar H (índice 7) - G (índice 6) y guardar en la columna I (índice 8)
##df.iloc[:, 8] = (df.iloc[:, 7] - df.iloc[:, 6]).dt.days


# Función para encontrar los nombres según la lógica dada
def obtener_nombres(fila):
    if pd.isna(fila.iloc[8]) or fila.iloc[8] <= 3:
        return None

    fecha_limite = fila.iloc[6] + timedelta(days=3)
    nombres_validos = []
    nombres_presentes = set()

    if pd.notna(fila.iloc[5]) and isinstance(fila.iloc[5], str) and "|" in fila.iloc[5]:
        nombres_fechas = fila.iloc[5].split(',')
        registros = []
        for item in nombres_fechas:
            partes = item.split('|')
            if len(partes) < 2:
                continue
            try:
                nombre = partes[0].strip()
                fecha = datetime.strptime(partes[1].strip(), "%m/%d/%Y %I:%M:%S %p")
                registros.append((fecha, nombre))
            except ValueError:
                continue
        registros.sort()
        for fecha, nombre in registros:
            if fecha > fecha_limite:
                nombres_validos.append(nombre)
            nombres_presentes.add(nombre)

    # Verificar nombres de la columna L (índice 11)
    if pd.notna(fila.iloc[11]):
        nombres_columna_L = [n.strip() for n in fila.iloc[11].split(',')]
        nombres_faltantes = [n for n in nombres_columna_L if n not in nombres_presentes]
        nombres_validos.extend(nombres_faltantes)

    return ', '.join(sorted(set(nombres_validos))) if nombres_validos else None


# Aplicar la función en cada fila
df.iloc[:, 9] = df.apply(obtener_nombres, axis=1)

# Guardar los cambios en un nuevo archivo
output_path = "resultado.xlsx"
df.to_excel(output_path, index=False, engine='openpyxl')
print(f"Proceso completado. Archivo guardado en {output_path}")