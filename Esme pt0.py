import pandas as pd
import re

# Cargar el archivo de Excel
file_path = "C:\\Users\\marco.avila\\Downloads\\Diciembre 2024.xlsx"  # Cambia esto por la ruta de tu archivo 
sheet_name = "in"  # Cambia esto si la hoja tiene otro nombre
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Eliminar columnas específicas
cols_to_drop = [2, 5, 7, 9]  # Columnas 3, 5, 6, y M (índices base 0)
df.drop(df.columns[cols_to_drop], axis=1, inplace=True)

# Insertar nuevas columnas entre B y C (índices base 0)
df.insert(2, "Solicitado por:", "")  # Nueva columna C
df.insert(3, "Gerencia Solicitante:", "")  # Nueva columna D

# Insertar 4 nuevas columnas entre G y H
df.insert(7, "Fecha de aprobación:", "")  # Nueva columna H
df.insert(8, "Tiempo de respuesta:", "")  # Nueva columna I
df.insert(9, "Retraso por:", "")  # Nueva columna J
df.insert(10, "Gerencia atrasada:", "")  # Nueva columna K

# Paso 0.07: Llenar columna C con datos de B
# Usar expresión regular para extraer nombres
patron = r'(?:Solicitado por|Solicitud por):\s*([^\n]*)'
df['Solicitado por:'] = df.iloc[:, 1].str.extract(patron, expand=False)

# Suponiendo que la columna C es la tercera (índice 2)
col_index = 2

# Función para truncar el texto a solo dos palabras
def truncar_texto(texto):
    if isinstance(texto, str):  # Verifica si es texto
        palabras = texto.split()
        return " ".join(palabras[:2])  # Toma solo las dos primeras palabras
    return texto  # Devuelve el valor original si no es texto

# Aplicar la función a la columna por índice
df.iloc[:, col_index] = df.iloc[:, col_index].apply(truncar_texto)

# Guardar el archivo modificado
df.to_excel("archivo_modificado22.xlsx", index=False)
