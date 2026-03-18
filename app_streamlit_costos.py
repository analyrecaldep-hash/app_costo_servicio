import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Cálculo de costo por servicio", layout="wide")

TARIFARIO = {
    'TIPO I': {
        'EFECTIVO': 97.00,
        'NO EFECTIVO': 48.50,
        'CANCELADO': 0.00,
    },
    'TIPO II': {
        'EFECTIVO': 125.00,
        'NO EFECTIVO': 62.50,
        'CANCELADO': 0.00,
    },
    'TIPO III': {
        'EFECTIVO': 1380.00,
        'NO EFECTIVO': 690.00,
        'CANCELADO': 0.00,
    },
    'TIPO III NEONATAL': {
        'EFECTIVO': 2760.00,
        'NO EFECTIVO': 1380.00,
        'CANCELADO': 0.00,
    },
}


def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ''
    return str(valor).strip().upper()


TIPO_UNIDAD_MAP = {
    'TIPO I': 'TIPO I',
    'TIPO 1': 'TIPO I',
    'I': 'TIPO I',
    'TIPO II': 'TIPO II',
    'TIPO 2': 'TIPO II',
    'II': 'TIPO II',
    'TIPO III': 'TIPO III',
    'TIPO 3': 'TIPO III',
    'III': 'TIPO III',
    'TIPO III NEONATAL': 'TIPO III NEONATAL',
    'TIPO 3 NEONATAL': 'TIPO III NEONATAL',
    'III NEONATAL': 'TIPO III NEONATAL',
    'NEONATAL': 'TIPO III NEONATAL',
}


RESULTADO_MAP = {
    'SI': 'EFECTIVO',
    'SÍ': 'EFECTIVO',
    'EFECTIVO': 'EFECTIVO',
    'ATENDIDO': 'EFECTIVO',
    'NO': 'NO EFECTIVO',
    'NO EFECTIVO': 'NO EFECTIVO',
    'NO_EFECTIVO': 'NO EFECTIVO',
    'NOEFECTIVO': 'NO EFECTIVO',
    'CANCELADO': 'CANCELADO',
    'CANCELADA': 'CANCELADO',
    'ANULADO': 'CANCELADO',
    'ANULADA': 'CANCELADO',
}


def homologar_tipo_unidad(valor: str) -> str:
    valor = normalizar_texto(valor)
    return TIPO_UNIDAD_MAP.get(valor, valor)



def homologar_resultado(valor: str) -> str:
    valor = normalizar_texto(valor)
    return RESULTADO_MAP.get(valor, valor)



def construir_tarifario_df() -> pd.DataFrame:
    filas = []
    for tipo_unidad, resultados in TARIFARIO.items():
        for resultado, costo in resultados.items():
            filas.append({
                'tipo_unidad_normalizado': tipo_unidad,
                'resultado_final': resultado,
                'Costo_servicio': costo,
            })
    return pd.DataFrame(filas)



def obtener_resultado_final(df: pd.DataFrame) -> pd.Series:
    efectivo = df['efectivo'].apply(homologar_resultado) if 'efectivo' in df.columns else pd.Series([''] * len(df))
    estado = df['estado'].apply(homologar_resultado) if 'estado' in df.columns else pd.Series([''] * len(df))

    resultado = []
    for e, s in zip(efectivo, estado):
        if e in ('EFECTIVO', 'NO EFECTIVO', 'CANCELADO'):
            resultado.append(e)
        elif s in ('EFECTIVO', 'NO EFECTIVO', 'CANCELADO'):
            resultado.append(s)
        else:
            resultado.append('')
    return pd.Series(resultado)



def procesar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if 'tipo_unidad' not in df.columns:
        raise ValueError("El archivo no contiene la columna 'tipo_unidad'.")
    if 'efectivo' not in df.columns and 'estado' not in df.columns:
        raise ValueError("El archivo debe contener 'efectivo' o 'estado'.")

    df = df.copy()
    df['tipo_unidad_normalizado'] = df['tipo_unidad'].apply(homologar_tipo_unidad)
    df['resultado_final'] = obtener_resultado_final(df)

    tarifario_df = construir_tarifario_df()
    df = df.merge(
        tarifario_df,
        how='left',
        on=['tipo_unidad_normalizado', 'resultado_final']
    )

    df['Costo_servicio'] = df['Costo_servicio'].fillna(0.0)

    def observacion(row):
        if row['tipo_unidad_normalizado'] not in TARIFARIO:
            return 'Tipo de unidad no reconocido'
        if row['resultado_final'] == '':
            return 'Resultado no reconocido'
        return 'OK'

    df['observacion_costo'] = df.apply(observacion, axis=1)
    return df



def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='resultado')
    return output.getvalue()


st.title("Cálculo web de costo por servicio")
st.caption("Sube el Excel del proveedor, calcula el costo por servicio y descarga el resultado.")

with st.expander("Tarifario aplicado", expanded=False):
    st.dataframe(construir_tarifario_df(), use_container_width=True, hide_index=True)

archivo = st.file_uploader("Sube un archivo Excel", type=["xlsx", "xls"])

if archivo is not None:
    try:
        df = pd.read_excel(archivo)
        st.success(f"Archivo cargado: {archivo.name}")
        st.write("Vista previa")
        st.dataframe(df.head(20), use_container_width=True)

        if st.button("Calcular costo_servicio", type="primary"):
            resultado = procesar_dataframe(df)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Registros", len(resultado))
            with col2:
                st.metric("Costo total", f"S/ {resultado['Costo_servicio'].sum():,.2f}")
            with col3:
                st.metric("Observaciones", int((resultado['observacion_costo'] != 'OK').sum()))

            st.subheader("Resultado")
            columnas_mostrar = [c for c in [
                'nro_solicitud', 'paciente', 'tipo_unidad', 'efectivo', 'estado',
                'tipo_unidad_normalizado', 'resultado_final', 'Costo_servicio', 'observacion_costo'
            ] if c in resultado.columns]
            st.dataframe(resultado[columnas_mostrar].head(50), use_container_width=True)

            excel_bytes = dataframe_to_excel_bytes(resultado)
            st.download_button(
                label="Descargar Excel con costo_servicio",
                data=excel_bytes,
                file_name="resultado_costo_servicio.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"No se pudo procesar el archivo: {e}")
