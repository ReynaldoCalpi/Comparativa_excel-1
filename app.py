import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Comparador y Sincronizador de Bases de Datos",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Comparador de Bases de Datos")
st.write(
    "Sube tus archivos de origen y destino para detectar registros nuevos, diferencias y elementos ausentes."
)

# 1. Widgets para la carga de archivos
col1, col2 = st.columns(2)
with col1:
  archivo_origen = st.file_uploader(
      "Sube la Base de **Origen** (Excel)", type=["xlsx", "xls"]
  )
with col2:
  archivo_destino = st.file_uploader(
      "Sube la Base de **Destino** (Excel)", type=["xlsx", "xls"]
  )

# Validar que ambos archivos estén cargados
if archivo_origen is not None and archivo_destino is not None:
  # Cargar los dataframes
  df_origen = pd.read_excel(archivo_origen)
  df_destino = pd.read_excel(archivo_destino)

  st.success("¡Archivos cargados correctamente!")

  # Mostrar una vista previa de los datos
  with st.expander("Ver vista previa de los datos"):
    c1, c2 = st.columns(2)
    with c1:
      st.write("**Origen (Primeras filas):**")
      st.dataframe(df_origen.head(3))
    with c2:
      st.write("**Destino (Primeras filas):**")
      st.dataframe(df_destino.head(3))

  st.divider()

  # Configuración de columnas clave y de auditoría
  st.subheader("⚙️ Configuración de la Comparación")
  columnas_comunes = list(
      set(df_origen.columns).intersection(set(df_destino.columns))
  )

  if len(columnas_comunes) > 0:
    col_k1, col_k2 = st.columns(2)

    with col_k1:
      llave = st.selectbox(
          "Selecciona la **Llave Primaria** (Identificador único común):",
          columnas_comunes,
      )

    with col_k2:
      columna_a_revisar = st.selectbox(
          "Selecciona una columna para auditar cambios de valor (Opcional):",
          columnas_comunes,
      )

    if st.button("Ejecutar Comparación", type="primary"):
      # 2. Realizar el merge externo para comparar
      comparacion = pd.merge(
          df_origen,
          df_destino,
          on=llave,
          how="outer",
          indicator=True,
          suffixes=("_origen", "_destino"),
      )

      # A. Registros nuevos (están en origen, faltan en destino)
      nuevos_registros = comparacion[
          comparacion["_merge"] == "left_only"
      ].copy()
      columnas_orig = [
          c for c in nuevos_registros.columns if not c.endswith("_destino")
      ]
      nuevos_registros = nuevos_registros[columnas_orig]
      nuevos_registros.columns = [
          c.replace("_origen", "") for c in nuevos_registros.columns
      ]

      # B. Registros eliminados / ausentes en origen (están en destino, faltan en origen)
      eliminados_registros = comparacion[
          comparacion["_merge"] == "right_only"
      ].copy()

      # C. Registros comunes para revisar posibles actualizaciones
      comunes = comparacion[comparacion["_merge"] == "both"].copy()

      if columna_a_revisar:
        cambios = comunes[
            comunes[f"{columna_a_revisar}_origen"]
            != comunes[f"{columna_a_revisar}_destino"]
        ]
      else:
        cambios = pd.DataFrame()

      # 3. Mostrar Resumen en Métricas
      st.divider()
      st.subheader("📈 Resultados de la Comparación")
      m1, m2, m3 = st.columns(3)
      m1.metric("Registros Nuevos (Insertar)", len(nuevos_registros))
      m2.metric(
          f"Con Diferencias en '{columna_a_revisar}' (Actualizar)", len(cambios)
      )
      m3.metric("Ausentes en Origen (Destino)", len(eliminados_registros))

      # Mostrar tablas interactivas de resultados
      tab1, tab2, tab3 = st.tabs(
          ["Nuevos Registros", "Actualizaciones", "Ausentes en Origen"]
      )

      with tab1:
        st.dataframe(nuevos_registros)
      with tab2:
        st.dataframe(cambios)
      with tab3:
        st.dataframe(eliminados_registros)

      # 4. Generar archivo Excel de salida para descarga
      import io

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        nuevos_registros.to_excel(writer, sheet_name="Nuevos", index=False)
        cambios.to_excel(writer, sheet_name="Actualizaciones", index=False)
        eliminados_registros.to_excel(
            writer, sheet_name="Ausentes_En_Origen", index=False
        )
      processed_data = output.getvalue()

      st.download_button(
          label="📥 Descargar Reporte de Diferencias en Excel",
          data=processed_data,
          file_name="reporte_diferencias_bases.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

  else:
    st.warning(
        "Los archivos seleccionados no tienen columnas en común para usar como"
        " llave."
    )
else:
  st.info(
      "👆 Por favor, sube ambos archivos de Excel en la parte superior para"
      " comenzar."
  )
