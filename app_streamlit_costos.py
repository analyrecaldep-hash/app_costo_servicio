import math
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Cálculo web de costo por servicio", layout="wide")

# =========================================================
# CONFIGURACION
# =========================================================
MINUTOS_LIBRES_ESPERA = 40
BLOQUE_ESPERA = 30

TARIFA_ESPERA = {
    "TIPO I": 45.0,
    "TIPO II": 45.0,
    "TIPO III": 200.0,
    "TIPO III NEONATAL": 200.0,
}

COSTO_SERVICIO = {
    ("TIPO I", "EFECTIVO"): 97.0,
    ("TIPO II", "EFECTIVO"): 125.0,
    ("TIPO III", "EFECTIVO"): 1380.0,
    ("TIPO III NEONATAL", "EFECTIVO"): 2760.0,
    ("TIPO I", "NO EFECTIVO"): 48.5,
    ("TIPO II", "NO EFECTIVO"): 62.5,
    ("TIPO III", "NO EFECTIVO"): 690.0,
    ("TIPO III NEONATAL", "NO EFECTIVO"): 1380.0,
}

# =========================================================
# UTILIDADES
# =========================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    valor = " ".join(valor.split())
    return valor

def safe_round(valor, dec=2):
    if pd.isna(valor):
        return np.nan
    return round(float(valor), dec)

def combinar_fecha_hora(fecha_col, hora_col):
    if pd.isna(fecha_col) or pd.isna(hora_col):
        return pd.NaT

    fecha = pd.to_datetime(fecha_col, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        return pd.NaT

    try:
        hora_dt = pd.to_datetime(hora_col, errors="coerce")
        if pd.notna(hora_dt):
            return fecha.normalize() + pd.Timedelta(
                hours=hora_dt.hour,
                minutes=hora_dt.minute,
                seconds=hora_dt.second
            )
    except Exception:
        pass

    try:
        hora_txt = str(hora_col).strip()
        if ":" in hora_txt:
            partes = hora_txt.split(":")
            hh = int(partes[0])
            mm = int(partes[1])
            ss = int(partes[2]) if len(partes) > 2 else 0
            return fecha.normalize() + pd.Timedelta(hours=hh, minutes=mm, seconds=ss)
    except Exception:
        pass

    return pd.NaT

def minutos_diff(inicio, fin):
    if pd.isna(inicio) or pd.isna(fin):
        return np.nan
    return (fin - inicio).total_seconds() / 60.0

def obtener_tarifa_espera(tipo_unidad):
    return TARIFA_ESPERA.get(normalizar_texto(tipo_unidad), 0.0)

def obtener_costo_servicio(tipo_unidad, efectivo):
    return COSTO_SERVICIO.get(
        (normalizar_texto(tipo_unidad), normalizar_texto(efectivo)),
        0.0
    )

def calcular_excedente_espera(minutos_espera):
    if pd.isna(minutos_espera) or minutos_espera <= MINUTOS_LIBRES_ESPERA:
        return 0.0
    return float(minutos_espera - MINUTOS_LIBRES_ESPERA)

def calcular_ocurrencias_espera(minutos_espera):
    excedente = calcular_excedente_espera(minutos_espera)
    if excedente <= 0:
        return 0
    return int(math.ceil(excedente / BLOQUE_ESPERA))

# =========================================================
# LOGICA DE NEGOCIO
# =========================================================
def calcular_tiempo_espera_origen(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")

    partida_origen = row.get("partida_origen")
    contacto_origen = row.get("contacto_paciente_origen")
    llegada_origen = row.get("llegada_origen")
    dt_programacion = row.get("dt_programacion")

    if (
        (motivo == "CITA" and sentido == "IDA" and modalidad in ["PROGRAMADA", "NO PROGRAMADA"]) or
        (motivo == "REFERENCIA" and sentido == "IDA" and modalidad in ["PROGRAMADA", "NO PROGRAMADA"]) or
        (motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA") or
        (motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA")
    ):
        return minutos_diff(contacto_origen, partida_origen)

    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "AMBAS(POR ERROR)":
        return minutos_diff(dt_programacion, partida_origen)

    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA"]:
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            if llegada_origen <= dt_programacion:
                return minutos_diff(dt_programacion, partida_origen)
            return minutos_diff(llegada_origen, partida_origen)

    return np.nan

def segunda_validacion_tiempo_espera_origen(row, minutos):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")

    contacto_origen = row.get("contacto_paciente_origen")
    llegada_origen = row.get("llegada_origen")
    dt_programacion = row.get("dt_programacion")
    dt_registro = row.get("dt_registro")

    if pd.isna(minutos):
        return np.nan

    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "AMBAS(POR ERROR)":
        if pd.notna(contacto_origen) and pd.notna(dt_programacion) and contacto_origen > dt_programacion:
            return 0.0

    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion) and llegada_origen > dt_programacion:
            return 0.0

    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_registro):
            dif = minutos_diff(dt_registro, llegada_origen)
            if pd.notna(dif) and dif > 30:
                return 0.0

    return minutos

def calcular_tiempo_espera_destino(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")

    llegada_destino = row.get("llegada_destino")
    hora_finalizacion = row.get("hora_finalizacion")
    dt_programacion = row.get("dt_programacion")

    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            if llegada_destino <= dt_programacion:
                return minutos_diff(dt_programacion, hora_finalizacion)
            return minutos_diff(llegada_destino, hora_finalizacion)

    if motivo == "CITA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(dt_programacion, hora_finalizacion)

    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "AMBAS(POR ERROR)":
        return minutos_diff(llegada_destino, hora_finalizacion)

    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            if llegada_destino <= dt_programacion:
                return minutos_diff(dt_programacion, hora_finalizacion)
            return minutos_diff(llegada_destino, hora_finalizacion)

    if motivo == "REFERENCIA" and sentido in ["IDA", "RETORNO"] and modalidad in ["NO PROGRAMADA", "PROGRAMADA"]:
        return minutos_diff(llegada_destino, hora_finalizacion)

    if motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    if motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    return np.nan

def segunda_validacion_tiempo_espera_destino(row, minutos):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")

    llegada_origen = row.get("llegada_origen")
    dt_programacion = row.get("dt_programacion")

    if pd.isna(minutos):
        return np.nan

    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            dif = abs(minutos_diff(llegada_origen, dt_programacion))
            if pd.notna(dif) and dif < 60:
                return 0.0

    return minutos

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
def procesar_archivo(df):
    df = df.copy()

    for col in ["sentido_traslado", "sede", "motivo_traslado", "modalidad", "estado", "efectivo", "tipo_unidad"]:
        if col in df.columns:
            df[col] = df[col].apply(normalizar_texto)

    if "motivo_traslado" in df.columns:
        df["motivo_traslado"] = df["motivo_traslado"].replace({
            "EMERGENCIAS": "EMERGENCIA",
            "REFERENCIAS": "REFERENCIA",
            "ALTAS": "ALTA"
        })

    if "modalidad" in df.columns:
        df["modalidad"] = df["modalidad"].replace({
            "PROGRAMADAS": "PROGRAMADA",
            "NO PROGRAMADO": "NO PROGRAMADA",
            "AMBAS(POR ERROR)": "AMBAS(POR ERROR)"
        })

    if "sentido_traslado" in df.columns:
        df["sentido_traslado"] = df["sentido_traslado"].replace({
            "RETORNO": "RETORNO",
            "IDA": "IDA"
        })

    columnas_datetime = [
        "salida_de_base",
        "llegada_origen",
        "contacto_paciente_origen",
        "partida_origen",
        "llegada_destino",
        "contacto_paciente_destino",
        "hora_finalizacion"
    ]

    for col in columnas_datetime:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["dt_registro"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_registro"), r.get("hora_registro")), axis=1
    )
    df["dt_programacion"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_programada"), r.get("hora_programada")), axis=1
    )

    def procesar_fila(row):
        estado = row.get("estado", "")
        tipo_unidad = row.get("tipo_unidad", "")
        efectivo = row.get("efectivo", "")

        tarifa_espera = obtener_tarifa_espera(tipo_unidad)
        costo_servicio = obtener_costo_servicio(tipo_unidad, efectivo)

        if estado == "CANCELADO":
            return pd.Series({
                "Costo_servicio": 0.0,
                "tarifa_espera": tarifa_espera,
                "min_espera_origen": np.nan,
                "minutos_excedentes_origen": 0.0,
                "ocurrencias_origen": 0,
                "Sobrecosto_tiempo_espera_origen": 0.0,
                "min_espera_destino": np.nan,
                "minutos_excedentes_destino": 0.0,
                "ocurrencias_destino": 0,
                "Sobrecosto_tiempo_espera_Destino": 0.0,
            })

        min_origen = segunda_validacion_tiempo_espera_origen(
            row, calcular_tiempo_espera_origen(row)
        )
        excedente_origen = calcular_excedente_espera(min_origen)
        ocurrencias_origen = calcular_ocurrencias_espera(min_origen)
        sobrecosto_origen = ocurrencias_origen * tarifa_espera

        min_destino = segunda_validacion_tiempo_espera_destino(
            row, calcular_tiempo_espera_destino(row)
        )
        excedente_destino = calcular_excedente_espera(min_destino)
        ocurrencias_destino = calcular_ocurrencias_espera(min_destino)
        sobrecosto_destino = ocurrencias_destino * tarifa_espera

        return pd.Series({
            "Costo_servicio": safe_round(costo_servicio),
            "tarifa_espera": tarifa_espera,
            "min_espera_origen": safe_round(min_origen),
            "minutos_excedentes_origen": safe_round(excedente_origen),
            "ocurrencias_origen": ocurrencias_origen,
            "Sobrecosto_tiempo_espera_origen": safe_round(sobrecosto_origen),
            "min_espera_destino": safe_round(min_destino),
            "minutos_excedentes_destino": safe_round(excedente_destino),
            "ocurrencias_destino": ocurrencias_destino,
            "Sobrecosto_tiempo_espera_Destino": safe_round(sobrecosto_destino),
        })

    resultado = df.apply(procesar_fila, axis=1)
    df_salida = pd.concat([df, resultado], axis=1)

    df_salida["tiempo_espera_total"] = (
        df_salida["min_espera_origen"].fillna(0) +
        df_salida["min_espera_destino"].fillna(0)
    ).round(2)

    df_salida["sobrecosto_total_espera"] = (
        df_salida["Sobrecosto_tiempo_espera_origen"].fillna(0) +
        df_salida["Sobrecosto_tiempo_espera_Destino"].fillna(0)
    ).round(2)

    df_salida["ocurrencias_total"] = (
        df_salida["ocurrencias_origen"].fillna(0) +
        df_salida["ocurrencias_destino"].fillna(0)
    ).astype(int)

    return df_salida

def exportar_excel(df_resultado):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, sheet_name="detalle", index=False)
    output.seek(0)
    return output

# =========================================================
# UI
# =========================================================
st.title("Cálculo web de costo por servicio")
st.caption("Sube el Excel del proveedor, calcula el costo por servicio y descarga el resultado.")

with st.expander("Tarifario aplicado"):
    st.markdown("""
    **Costo servicio**
    - Tipo I + Efectivo: 97
    - Tipo II + Efectivo: 125
    - Tipo III + Efectivo: 1380
    - Tipo III Neonatal + Efectivo: 2760
    - Tipo I + No efectivo: 48.5
    - Tipo II + No efectivo: 62.5
    - Tipo III + No efectivo: 690
    - Tipo III Neonatal + No efectivo: 1380

    **Tiempo de espera**
    - Tipo I: 45
    - Tipo II: 45
    - Tipo III: 200
    - Tipo III Neonatal: 200

    **Regla de espera**
    - 0 a 40 min: no cobra
    - desde 41 min: cobra por bloques de 30 min
    """)

archivo = st.file_uploader("Sube un archivo Excel", type=["xlsx", "xls"])

if archivo is not None:
    try:
        df_base = pd.read_excel(archivo)
        df_resultado = procesar_archivo(df_base)

        st.success("Archivo procesado correctamente.")

        # =========================================
        # FILTROS TIPO POWER BI
        # =========================================
        st.sidebar.header("Filtros")

        estados_disponibles = sorted(df_resultado["estado"].dropna().unique())
        estados_default = [x for x in ["POR VALIDAR", "VALIDADO"] if x in estados_disponibles]
        if not estados_default:
            estados_default = estados_disponibles

        estados = st.sidebar.multiselect(
            "Estado",
            options=estados_disponibles,
            default=estados_default
        )

        sedes = st.sidebar.multiselect(
            "Sede",
            options=sorted(df_resultado["sede"].dropna().unique()),
            default=sorted(df_resultado["sede"].dropna().unique())
        )

        tipos = st.sidebar.multiselect(
            "Tipo unidad",
            options=sorted(df_resultado["tipo_unidad"].dropna().unique()),
            default=sorted(df_resultado["tipo_unidad"].dropna().unique())
        )

        df_filtrado = df_resultado[
            (df_resultado["estado"].isin(estados)) &
            (df_resultado["sede"].isin(sedes)) &
            (df_resultado["tipo_unidad"].isin(tipos))
        ].copy()

        # =========================================
        # MÉTRICAS
        # =========================================
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros", len(df_filtrado))
        c2.metric("Costo servicio total", f"{df_filtrado['Costo_servicio'].sum():,.2f}")
        c3.metric("Sobrecosto total espera", f"{df_filtrado['sobrecosto_total_espera'].sum():,.2f}")

        # =========================================
        # VISTA PREVIA
        # =========================================
        st.subheader("Vista previa")
        st.dataframe(df_filtrado.head(20), use_container_width=True)

        # =========================================
        # RESUMEN POR TIPO DE UNIDAD
        # =========================================
        st.subheader("Resumen por tipo de unidad")
        resumen_tipo = df_filtrado.groupby("tipo_unidad", dropna=False).agg(
            Cantidad=("tipo_unidad", "size"),
            Costo_servicio=("Costo_servicio", "sum"),
            sobrecosto_total_espera=("sobrecosto_total_espera", "sum")
        ).reset_index()

        fila_total_tipo = {
            "tipo_unidad": "TOTAL",
            "Cantidad": resumen_tipo["Cantidad"].sum(),
            "Costo_servicio": resumen_tipo["Costo_servicio"].sum(),
            "sobrecosto_total_espera": resumen_tipo["sobrecosto_total_espera"].sum()
        }

        resumen_tipo = pd.concat(
            [resumen_tipo, pd.DataFrame([fila_total_tipo])],
            ignore_index=True
        )

        resumen_tipo_formateado = resumen_tipo.copy()
        for col in ["Cantidad", "Costo_servicio", "sobrecosto_total_espera"]:
            resumen_tipo_formateado[col] = resumen_tipo_formateado[col].apply(
                lambda x: f"{x:,.0f}" if pd.notna(x) else ""
            )

        st.dataframe(resumen_tipo_formateado, use_container_width=True)

        # =========================================
        # RESUMEN POR SEDE
        # =========================================
        st.subheader("Resumen por sede")
        resumen_sede = df_filtrado.groupby("sede", dropna=False).agg(
            Cantidad=("sede", "size"),
            Costo_servicio=("Costo_servicio", "sum"),
            sobrecosto_total_espera=("sobrecosto_total_espera", "sum")
        ).reset_index()

        fila_total_sede = {
            "sede": "TOTAL",
            "Cantidad": resumen_sede["Cantidad"].sum(),
            "Costo_servicio": resumen_sede["Costo_servicio"].sum(),
            "sobrecosto_total_espera": resumen_sede["sobrecosto_total_espera"].sum()
        }

        resumen_sede = pd.concat(
            [resumen_sede, pd.DataFrame([fila_total_sede])],
            ignore_index=True
        )

        resumen_sede_formateado = resumen_sede.copy()
        for col in ["Cantidad", "Costo_servicio", "sobrecosto_total_espera"]:
            resumen_sede_formateado[col] = resumen_sede_formateado[col].apply(
                lambda x: f"{x:,.0f}" if pd.notna(x) else ""
            )

        st.dataframe(resumen_sede_formateado, use_container_width=True)
        # =========================================
        # TABLA CRUZADA
        # =========================================
        st.subheader("Tabla cruzada por tipo y sede")

        cantidad = pd.pivot_table(
            df_filtrado,
            index="tipo_unidad",
            columns="sede",
            values="Costo_servicio",
            aggfunc="count",
            fill_value=0
        )

        tarifa = pd.pivot_table(
            df_filtrado,
            index="tipo_unidad",
            columns="sede",
            values="Costo_servicio",
            aggfunc="sum",
            fill_value=0
        )

        partes = []
        for sede in cantidad.columns:
            temp = pd.DataFrame({
                (sede, "Cantidad"): cantidad[sede],
                (sede, "Tarifa Base."): tarifa[sede]
            })
            partes.append(temp)

        tabla_cruzada = pd.concat(partes, axis=1)

        tabla_cruzada[("Total", "Cantidad")] = cantidad.sum(axis=1)
        tabla_cruzada[("Total", "Tarifa Base.")] = tarifa.sum(axis=1)

        tabla_cruzada.loc["Total general"] = tabla_cruzada.sum()

        tabla_cruzada = tabla_cruzada.reset_index().rename(columns={"tipo_unidad": "Etiquetas de fila"})

        st.dataframe(tabla_cruzada, use_container_width=True)
        # =========================================
        # DESCARGA
        # =========================================
        excel_bytes = exportar_excel(df_resultado)
        st.download_button(
            "Descargar resultado",
            data=excel_bytes,
            file_name="resultado_tiempo_espera.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
