import pandas as pd

# 1. Cargar las bases de datos (pueden ser CSV, Excel, o consultas a SQL)
# Supongamos que leemos dos archivos o tablas
df_origen = pd.read_excel("base_origen.xlsx")
df_destino = pd.read_excel("base_destino.xlsx")

# Definir la llave primaria o identificador único común
llave = "Codigo"  # Reemplaza con tu campo clave (ej. 'NIT', 'ID', 'Factura')

# 2. Realizar un merge externo para comparar ambos DataFrames
# suffixes nos ayuda a distinguir columnas de origen vs destino si tienen el mismo nombre
comparacion = pd.merge(
    df_origen, 
    df_destino, 
    on=llave, 
    how="outer", 
    indicator=True, 
    suffixes=('_origen', '_destino')
)

# 3. Clasificar los resultados según el indicador de Pandas ('both', 'left_only', 'right_only')

# A. Nuevos registros (están en origen, faltan en destino)
nuevos_registros = comparacion[comparacion['_merge'] == 'left_only'].copy()
# Limpiamos las columnas para quedarnos con los datos del origen
columnas_origen = [c for c in nuevos_registros.columns if not c.endswith('_destino')]
nuevos_registros = nuevos_registros[columnas_origen]
nuevos_registros.columns = [c.replace('_origen', '') for c in nuevos_registros.columns]

# B. Registros que ya no están (están en destino, faltan en origen)
eliminados_registros = comparacion[comparacion['_merge'] == 'right_only'].copy()

# C. Registros comunes para revisar posibles actualizaciones
comunes = comparacion[comparacion['_merge'] == 'both'].copy()

# 4. Detectar cambios en campos específicos dentro de los registros comunes
# Ejemplo: comparar la columna 'Precio' o 'Saldo'
columna_a_revisar = 'Saldo'  # Reemplaza por la columna que deseas auditar
cambios = comunes[comunes[f'{columna_a_revisar}_origen'] != comunes[f'{columna_a_revisar}_destino']]

# 5. Mostrar un resumen de lo encontrado
print(f"--- RESUMEN DE COMPARACIÓN ---")
print(f"Registros nuevos a insertar: {len(nuevos_registros)}")
print(f"Registros con diferencias a actualizar: {len(cambios)}")
print(f"Registros ausentes en origen (destino): {len(eliminados_registros)}")

# Opcional: Exportar los resultados a un Excel para revisión visual
with pd.ExcelWriter('reporte_diferencias.xlsx') as writer:
    nuevos_registros.to_excel(writer, sheet_name='Nuevos', index=False)
    cambios.to_excel(writer, sheet_name='Actualizaciones', index=False)
    eliminados_registros.to_excel(writer, sheet_name='Eliminados', index=False)