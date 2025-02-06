import streamlit as st
import numpy as np
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO  # Importación añadida

# Configuración de la aplicación Streamlit
st.title("Procesamiento de Archivos Excel")

# Cargar archivo de Excel
uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])

if uploaded_file:
    sheet_name = st.text_input("Nombre de la hoja", "in")
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # Eliminar columnas específicas
    cols_to_drop = [2, 5, 7, 9]
    df.drop(df.columns[cols_to_drop], axis=1, inplace=True)

    # Insertar nuevas columnas
    df.insert(2, "Solicitado por:", "")
    df.insert(3, "Gerencia Solicitante:", "")
    df.insert(7, "Fecha de aprobación:", "")
    df.insert(8, "Tiempo de respuesta:", "")
    df.insert(9, "Retraso por:", "")
    df.insert(10, "Gerencia atrasada:", "")

    # Extraer nombres usando expresión regular
    patron = r'(?:Solicitado por|Solicitud por):\s*([^\n]*)'
    df['Solicitado por:'] = df.iloc[:, 1].str.extract(patron, expand=False)

    # Función para truncar el texto a solo dos palabras
    def truncar_texto(texto):
        if isinstance(texto, str):
            palabras = texto.split()
            return " ".join(palabras[:2])
        return texto

    df.iloc[:, 2] = df.iloc[:, 2].apply(truncar_texto)

    # Limpiar caracteres no alfanuméricos y convertir fechas
    df.iloc[:, 6] = df.iloc[:, 6].astype(str).str.replace(r'[^\w\s]', '', regex=True).str.strip()
    df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6], format='%d%m%Y', errors='coerce')
    df.iloc[:, 6] = df.iloc[:, 6].apply(lambda x: x.strftime('%m%d%Y') if pd.notnull(x) else '')
    df.iloc[:, 6] = df.iloc[:, 6].astype(str).apply(lambda x: f'{x[:2]}/{x[2:4]}/{x[4:]}')
    df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6], format='%m/%d/%Y', errors='coerce')

    # Lista de días inhábiles
    dias_inhabiles = {
        datetime(2025, 1, 1), datetime(2025, 2, 3), datetime(2025, 3, 17),
        datetime(2025, 4, 18), datetime(2025, 5, 1), datetime(2025, 9, 15),
        datetime(2025, 11, 17), datetime(2025, 12, 24), datetime(2025, 12, 25), datetime(2025, 12, 31)
    }

    # Función para sumar días hábiles
    def sumar_dias_habiles(fecha, dias):
        while dias > 0:
            fecha += timedelta(days=1)
            if fecha.weekday() < 5 and fecha not in dias_inhabiles:
                dias -= 1
        return fecha

    # Calcular "Límite Real"
    df.insert(12, "Límite Real", df.iloc[:, 6].apply(lambda x: sumar_dias_habiles(x, 3) if pd.notna(x) else None))

    # Función para extraer la última fecha
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

    # Aplicar extracción de fechas y cálculos de diferencias
    df.iloc[:, 7] = df.iloc[:, 5].apply(extraer_ultima_fecha)
    df.iloc[:, 7] = pd.to_datetime(df.iloc[:, 7])
    df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6])

    # Calcular la diferencia de días considerando valores nulos
    df.iloc[:, 8] = df.apply(lambda row: (row[7] - row[12]).days if pd.notnull(row[7]) and pd.notnull(row[12]) else 0, axis=1)

    # Reemplazar valores negativos por 0
    df.iloc[:, 8] = df.iloc[:, 8].apply(lambda x: max(x, 0))

    # Función para obtener nombres retrasados
    def obtener_nombres(fila):
        if pd.isna(fila.iloc[8]) or fila.iloc[8] <= .99:
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

        if pd.notna(fila.iloc[11]):
            nombres_columna_L = [n.strip() for n in fila.iloc[11].split(',')]
            nombres_faltantes = [n for n in nombres_columna_L if n not in nombres_presentes]
            nombres_validos.extend(nombres_faltantes)

        return ', '.join(sorted(set(nombres_validos))) if nombres_validos else None

    # Aplicar la función en cada fila
    df.iloc[:, 9] = df.apply(obtener_nombres, axis=1)

    # Mostrar resultados en Streamlit
    st.write("Vista previa de los datos procesados:")
    st.dataframe(df)

    # Descargar el resultado final usando BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    st.download_button("Descargar resultado", output, file_name="resultado_final.xlsx")