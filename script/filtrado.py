import pandas as pd

# Leer el archivo Excel
path = r"C:\Users\samue\Desktop\AlgoritmoFiltrado\input\hojaDatos.xlsx"
df = pd.read_excel(path,header=None)

# Extraer la primera fila para usar como nombres de columnas
column_names = df.iloc[0, 0].split(';')

# Separar los datos de la única columna en varias columnas usando el delimitador `;`
data = df.iloc[1:, 0].str.split(';', expand=True)

# Asignar los nombres de columnas
data.columns = column_names

# Resetear el índice del DataFrame resultante
data.reset_index(drop=True, inplace=True)

# Filtrar los registros donde en las columnas 7, 8, 9 o 10 contienen la palabra "industrial"
filtered_data = data[
    data.iloc[:, 6].str.contains('industrial', case=False, na=False) |
    data.iloc[:, 7].str.contains('industrial', case=False, na=False) |
    data.iloc[:, 8].str.contains('industrial', case=False, na=False) |
    data.iloc[:, 9].str.contains('industrial', case=False, na=False)
]

# Ruta del archivo de salida
output_path = r"C:\Users\samue\Desktop\AlgoritmoFiltrado\output\resultado.xlsx"

# Guardar los registros filtrados en un nuevo archivo Excel
filtered_data.to_excel(output_path, index=False)

print(f"Los primeros 3 registros se han guardado en {output_path}")