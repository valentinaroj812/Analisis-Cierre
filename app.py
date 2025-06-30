
import streamlit as st
import pandas as pd

st.title("Análisis de Cierres por Oficina")

uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Buscar la fila donde está el encabezado
    header_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains('Fecha Cierre').any(), axis=1)].index[0]

    # Leer el archivo nuevamente con la fila correcta como encabezado
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row)

    if 'Fecha Cierre' not in df.columns:
        st.error("No se encontró la columna 'Fecha Cierre'. Verifica que tu archivo tenga ese encabezado.")
        st.stop()

    # Procesamiento de datos
    df['Fecha Cierre'] = pd.to_datetime(df['Fecha Cierre'], errors='coerce')
    df['Precio Cierre'] = df['Precio Cierre'].replace('[\$,]', '', regex=True).astype(float)
    df = df.dropna(subset=['Fecha Cierre', 'Precio Cierre', 'Asesor Colocador'])

    df['Mes'] = df['Fecha Cierre'].dt.month
    df['Año'] = df['Fecha Cierre'].dt.year

    año = st.selectbox("Selecciona el año", sorted(df['Año'].dropna().unique(), reverse=True))
    mes = st.selectbox("Selecciona el mes", sorted(df[df['Año'] == año]['Mes'].dropna().unique()))

    df_filtrado = df[(df['Año'] == año) & (df['Mes'] == mes)]

    resumen = df_filtrado.groupby('Asesor Colocador').agg(
        Cantidad_Cierres=('Precio Cierre', 'count'),
        Monto_Total=('Precio Cierre', 'sum')
    ).reset_index()

    resumen['Promedio_por_Cierre'] = resumen['Monto_Total'] / resumen['Cantidad_Cierres']
    resumen = resumen.sort_values(by='Monto_Total', ascending=False)

    st.subheader("Resumen por Oficina/Asesor")
    st.dataframe(resumen)

    st.subheader("Detalle de Cierres")
    st.dataframe(df_filtrado[['Fecha Cierre', 'Asesor Colocador', 'Precio Cierre']])
