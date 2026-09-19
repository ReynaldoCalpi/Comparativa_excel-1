import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Comparador y Sincronizador de Placas",
    page_icon="🚗",
    layout="wide",
)

st.title("🚗 Comparador de Equipos de Transporte por Placa")
st.write(
    "Sube tus archivos, selecciona la columna de placa y el sistema limpiará"
    " automáticamente los prefijos (como C- o RE-) para comparar los últimos 5"
    " caracteres."
)

col1, col2 = st.columns(2)
with col1:
  archivo_origen = st.file_uploader(
      "Sube la Base de **Origen** (Excel)", type=["xlsx", "xls"], key="orig"
  )
with col2:
  archivo_destino = st.file_uploader(
      "Sube la Base de **Destino** (Excel)", type=["xlsx", "xls"], key="dest"
  )

if archivo_origen is not None and archivo_destino is not None:
  df_origen = pd.read_excel(archivo_origen)
  df_destino = pd.read_excel(archivo_destino)

  st.success("¡Archivos cargados correctamente!")

  with st.expander("Ver vista previa de los datos"):
    c1, c2 = st.columns(2)
    with c1:
      st.write("**Origen (Primeras filas):**")
      st.dataframe(df_origen.head(3))
    with c2:
      st.write("**Destino (Primeras filas):**")
      st.dataframe(df_destino.head(3))

  st.divider()
  st.subheader("⚙️ Configuración de la Llave de Placas")

  columnas_origen = list(df_origen.columns)
  columnas_destino = list(df_destino.columns)

  col_k1, col_k2 = st.columns(2)
  with col_k1:
    llave_origen = st.selectbox(
        "Columna de Placa en **Origen**:", columnas_origen
    )
  with col_k2:
    llave_destino = st.selectbox(
        "Columna de Placa en **Destino**:", columnas_destino
    )

  columnas_comunes = list(
      set(df_origen.columns).intersection(set(df_destino.columns))
  )
  columna_a_revisar = st.selectbox(
      "Selecciona una columna adicional para auditar cambios (Opcional):",
      ["Ninguna"] + columnas_comunes,
  )

  if st.button("Ejecutar Comparación por Placas", type="primary"):
    df_o = df_origen.copy()
    df_d = df_destino.copy()

    # Limpieza: extraer los últimos 5 caracteres
    df_o["_llave_limpia"] = (
        df_o[llave_origen].astype(str).str.strip().str[-5:]
    )
    df_d["_llave_limpia"] = (
        df_d[llave_destino].astype(str).str.strip().str[-5:]
    )

    comparacion = pd.merge(
        df_o,
        df_d,
        on="_llave_limpia",
        how="outer",
        indicator=True,
        suffixes=("_origen", "_destino"),
    )

    # A. Nuevos registros (están en origen, faltan en destino)
    nuevos_registros = comparacion[comparacion["_merge"] == "left_only"].copy()
    cols_orig = [
        c
        for c in nuevos_registros.columns
        if not c.endswith("_destino") and c != "_llave_limpia"
    ]
    nuevos_registros = nuevos_registros[cols_orig]
    nuevos_registros.columns = [
        c.replace("_origen", "") for c in nuevos_registros.columns
    ]

    # B. Registros ausentes en origen (están en destino, faltan en origen)
    eliminados_registros = comparacion[
        comparacion["_merge"] == "right_only"
    ].copy()
    cols_dest = [
        c
        for c in eliminados_registros.columns
        if not c.endswith("_origen") and c != "_llave_limpia"
    ]
    eliminados_registros = eliminados_registros[cols_dest]
    eliminados_registros.columns = [
        c.replace("_destino", "") for c in eliminados_registros.columns
    ]

    # C. Registros que SÍ hacen match (comunes)
    comunes = comparacion[comparacion["_merge"] == "both"].copy()

    # D. Auditoría de cambios dentro de los comunes
    cambios = pd.DataFrame()
    if columna_a_revisar != "Ninguna":
      col_orig_rev = f"{columna_a_revisar}_origen"
      col_dest_rev = f"{columna_a_revisar}_destino"
      if col_orig_rev in comunes.columns and col_dest_rev in comunes.columns:
        cambios = comunes[comunes[col_orig_rev] != comunes[col_dest_rev]]

    # Métricas de Resultados
    st.divider()
    st.subheader("📈 Resultados de la Comparación de Placas")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Placas Nuevas (Insertar)", len(nuevos_registros))
    m2.metric("Coincidencias (Match)", len(comunes))
    if columna_a_revisar != "Ninguna":
      m3.metric(f"Con Diferencias", len(cambios))
    else:
      m3.metric("Con Diferencias", 0)
    m4.metric("Placas Ausentes en Origen", len(eliminados_registros))

    # Pestañas con detalle (incluyendo la de coincidencias)
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Nuevos Registros", "Coincidencias (Match)", "Actualizaciones / Diferencias", "Ausentes en Origen"]
    )
    with tab1:
      st.dataframe(nuevos_registros)
    with tab2:
      st.dataframe(comunes)
    with tab3:
      st.dataframe(cambios)
    with tab4:
      st.dataframe(eliminados_registros)

    # Generar archivo Excel de salida con todas las hojas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      nuevos_registros.to_excel(writer, sheet_name="Nuevos", index=False)
      comunes.to_excel(writer, sheet_name="Coincidencias_Match", index=False)
      if not cambios.empty:
        cambios.to_excel(writer, sheet_name="Actualizaciones", index=False)
      eliminados_registros.to_excel(
          writer, sheet_name="Ausentes_En_Origen", index=False
      )
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Reporte Completo en Excel (Incluye Match)",
        data=processed_data,
        file_name="reporte_placas_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
  st.info("👆 Por favor, sube ambos archivos de Excel para comenzar.")
