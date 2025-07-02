
import streamlit as st
import pandas as pd
import altair as alt

st.title("Análisis de Cierres por Oficina")

uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Buscar la fila donde está el encabezado
    header_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains('Fecha Cierre').any(), axis=1)].index[0]

    # Leer el archivo nuevamente con la fila correcta como encabezado
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)

    if 'Fecha Cierre' not in df.columns or 'Precio Cierre' not in df.columns:
        st.error("No se encontraron las columnas necesarias. Verifica que el archivo tenga 'Fecha Cierre' y 'Precio Cierre'.")
        st.stop()

    # Procesamiento de datos
    df['Fecha Cierre'] = pd.to_datetime(df['Fecha Cierre'], errors='coerce')
    df['Precio Cierre'] = df['Precio Cierre'].replace('[\$,]', '', regex=True).astype(float)
    df = df.dropna(subset=['Fecha Cierre', 'Precio Cierre'])

    df['Mes'] = df['Fecha Cierre'].dt.month
    df['Año'] = df['Fecha Cierre'].dt.year

    st.sidebar.title("Filtros")
    año = st.sidebar.selectbox("Selecciona el año", sorted(df['Año'].dropna().unique(), reverse=True))
    mes = st.sidebar.selectbox("Selecciona el mes", sorted(df[df['Año'] == año]['Mes'].dropna().unique()))
    modo = st.sidebar.selectbox("Analizar por", ['Asesor Colocador', 'Asesor Captador'])

    df_filtrado = df[(df['Año'] == año) & (df['Mes'] == mes)]
    df_filtrado = df_filtrado.dropna(subset=[modo])

    resumen = df_filtrado.groupby(modo).agg(
        Cantidad_Cierres=('Precio Cierre', 'count'),
        Monto_Total=('Precio Cierre', 'sum')
    ).reset_index()

    resumen['Promedio_por_Cierre'] = resumen['Monto_Total'] / resumen['Cantidad_Cierres']
    resumen = resumen.sort_values(by='Monto_Total', ascending=False)

    st.subheader(f"Resumen por {modo}")
    st.dataframe(resumen)

    chart = alt.Chart(resumen).mark_bar().encode(
        x=alt.X(modo, sort='-y'),
        y='Monto_Total',
        tooltip=[modo, 'Cantidad_Cierres', 'Monto_Total']
    ).properties(width=700, height=400).interactive()

    st.altair_chart(chart)

    st.subheader("Detalle de Cierres")
    st.dataframe(df_filtrado[['Fecha Cierre', 'Asesor Colocador', 'Asesor Captador', 'Precio Cierre']])

    # Botón para descargar Excel
    import io
    output = io.BytesIO()
    df_filtrado.to_excel(output, index=False)
    st.download_button("📥 Descargar detalle en Excel", data=output.getvalue(), file_name="cierres_filtrados.xlsx")
