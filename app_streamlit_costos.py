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

# Penalidades fijas por tipo
PENALIDAD_TARIFA = {
    "TIPO I": 32.33,
    "TIPO II": 41.67,
    "TIPO III": 460.00,
    "TIPO III NEONATAL": 920.00,
}

PENALIDAD_PERDIDA_CITA = 535.00

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

def obtener_penalidad_tarifa(tipo_unidad):
    return PENALIDAD_TARIFA.get(normalizar_texto(tipo_unidad), 0.0)

def calcular_excedente_espera(minutos_espera):
    if pd.isna(minutos_espera) or minutos_espera <= MINUTOS_LIBRES_ESPERA:
        return 0.0
    return float(minutos_espera - MINUTOS_LIBRES_ESPERA)

def calcular_ocurrencias_espera(minutos_espera):
    excedente = calcular_excedente_espera(minutos_espera)
    if excedente <= 0:
        return 0
    return int(math.ceil(excedente / BLOQUE_ESPERA))

def calcular_bloques_penalidad(minutos, bloque=30):
    if pd.isna(minutos) or minutos <= 0:
        return 0
    return int(math.ceil(minutos / bloque))

# =========================================================
# LOGICA DE ESPERA
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
# LOGICA DE PENALIDADES
# =========================================================
def calcular_penalidad_origen(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")
    tipo_unidad = row.get("tipo_unidad", "")

    contacto_origen = row.get("contacto_paciente_origen")
    dt_programacion = row.get("dt_programacion")

    tarifa_penalidad = obtener_penalidad_tarifa(tipo_unidad)

    diferencia_min = np.nan
    minutos_penalidad = 0
    bloques_penalidad = 0
    penalidad_origen = 0.0

    # CITA / RETORNO / PROGRAMADA o NO PROGRAMADA
    # Regla: Hora de Programación - Hora de contacto con paciente
    if motivo == "CITA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA"]:
        if pd.notna(contacto_origen) and pd.notna(dt_programacion):
            diferencia_control = minutos_diff(contacto_origen, dt_programacion)

            if contacto_origen <= dt_programacion:
                diferencia_min = safe_round(diferencia_control)
                minutos_penalidad = 0
            else:
                diferencia_min = safe_round(minutos_diff(dt_programacion, contacto_origen))
                minutos_penalidad = diferencia_min if pd.notna(diferencia_min) and diferencia_min > 0 else 0

            bloques_penalidad = calcular_bloques_penalidad(minutos_penalidad, 30)
            penalidad_origen = bloques_penalidad * tarifa_penalidad

    return pd.Series({
        "diferencia_penalidad_origen": diferencia_min,
        "min_penalidad_origen": safe_round(minutos_penalidad),
        "bloques_penalidad_origen": bloques_penalidad,
        "penalidad_origen": safe_round(penalidad_origen)
    })

def calcular_penalidad_destino(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")
    tipo_unidad = row.get("tipo_unidad", "")

    llegada_destino = row.get("llegada_destino")
    llegada_origen = row.get("llegada_origen")
    dt_programacion = row.get("dt_programacion")

    tarifa_penalidad = obtener_penalidad_tarifa(tipo_unidad)

    diferencia_min = np.nan
    minutos_penalidad = 0
    bloques_penalidad = 0
    penalidad_destino = 0.0
    perdida_cita_flag = 0
    perdida_cita_monto = 0.0

    # REFERENCIA / IDA / PROGRAMADA
    # Regla: Hora de Llegada a Destino - Hora de Programación
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":

        # Si no hubo llegada a destino => pérdida de cita
        if pd.isna(llegada_destino):
            return pd.Series({
                "diferencia_penalidad_destino": np.nan,
                "min_penalidad_destino": 0,
                "bloques_penalidad_destino": 0,
                "penalidad_destino": 0.0,
                "perdida_cita_flag": 1,
                "perdida_cita_monto": PENALIDAD_PERDIDA_CITA
            })

        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            diferencia_min = minutos_diff(dt_programacion, llegada_destino)
            minutos_penalidad = diferencia_min if pd.notna(diferencia_min) and diferencia_min > 0 else 0

            # Segunda validación:
            # Hora de Programación - Hora de Llegada a Origen >= 90
            if pd.notna(llegada_origen):
                anticipacion_origen = minutos_diff(llegada_origen, dt_programacion)
                if pd.notna(anticipacion_origen) and anticipacion_origen >= 90:
                    minutos_penalidad = 0

            bloques_penalidad = calcular_bloques_penalidad(minutos_penalidad, 30)
            penalidad_destino = bloques_penalidad * tarifa_penalidad

    return pd.Series({
        "diferencia_penalidad_destino": safe_round(diferencia_min),
        "min_penalidad_destino": safe_round(minutos_penalidad),
        "bloques_penalidad_destino": bloques_penalidad,
        "penalidad_destino": safe_round(penalidad_destino),
        "perdida_cita_flag": perdida_cita_flag,
        "perdida_cita_monto": safe_round(perdida_cita_monto)
    })

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

    penalidad_origen = df_salida.apply(calcular_penalidad_origen, axis=1)
    penalidad_destino = df_salida.apply(calcular_penalidad_destino, axis=1)

    df_salida = pd.concat([df_salida, penalidad_origen, penalidad_destino], axis=1)

    df_salida["penalidad_total"] = (
        df_salida["penalidad_origen"].fillna(0) +
        df_salida["penalidad_destino"].fillna(0) +
        df_salida["perdida_cita_monto"].fillna(0)
    ).round(2)

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
st.caption("Sube el Excel del proveedor, calcula el costo por servicio, penalidades y descarga el resultado.")

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

    **Penalidad por bloque**
    - Tipo I: 32.33
    - Tipo II: 41.67
    - Tipo III: 460.00
    - Tipo III Neonatal: 920.00

    **Pérdida de cita**
    - 10% UIT: 535.00

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
        # FILTROS
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
        # MÉTRICAS PRINCIPALES
        # =========================================
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros", len(df_filtrado))
        c2.metric("Costo servicio total", f"S/ {df_filtrado['Costo_servicio'].sum():,.2f}")
        c3.metric("Sobrecosto total espera", f"S/ {df_filtrado['sobrecosto_total_espera'].sum():,.2f}")

        c4, c5, c6 = st.columns(3)
        c4.metric("Penalidad origen", f"S/ {df_filtrado['penalidad_origen'].sum():,.2f}")
        c5.metric("Penalidad destino", f"S/ {df_filtrado['penalidad_destino'].sum():,.2f}")
        c6.metric("Pérdida de cita", f"S/ {df_filtrado['perdida_cita_monto'].sum():,.2f}")

        c7, = st.columns(1)
        c7.metric("Penalidad total", f"S/ {df_filtrado['penalidad_total'].sum():,.2f}")

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
            sobrecosto_total_espera=("sobrecosto_total_espera", "sum"),
            penalidad_origen=("penalidad_origen", "sum"),
            penalidad_destino=("penalidad_destino", "sum"),
            perdida_cita=("perdida_cita_monto", "sum"),
            penalidad_total=("penalidad_total", "sum")
        ).reset_index()

        fila_total_tipo = {
            "tipo_unidad": "TOTAL",
            "Cantidad": resumen_tipo["Cantidad"].sum(),
            "Costo_servicio": resumen_tipo["Costo_servicio"].sum(),
            "sobrecosto_total_espera": resumen_tipo["sobrecosto_total_espera"].sum(),
            "penalidad_origen": resumen_tipo["penalidad_origen"].sum(),
            "penalidad_destino": resumen_tipo["penalidad_destino"].sum(),
            "perdida_cita": resumen_tipo["perdida_cita"].sum(),
            "penalidad_total": resumen_tipo["penalidad_total"].sum()
        }

        resumen_tipo = pd.concat(
            [resumen_tipo, pd.DataFrame([fila_total_tipo])],
            ignore_index=True
        )

        resumen_tipo_formateado = resumen_tipo.copy()
        for col in ["Cantidad"]:
            resumen_tipo_formateado[col] = resumen_tipo_formateado[col].apply(
                lambda x: f"{x:,.0f}" if pd.notna(x) else ""
            )
        for col in ["Costo_servicio", "sobrecosto_total_espera", "penalidad_origen", "penalidad_destino", "perdida_cita", "penalidad_total"]:
            resumen_tipo_formateado[col] = resumen_tipo_formateado[col].apply(
                lambda x: f"S/ {x:,.2f}" if pd.notna(x) else ""
            )

        st.dataframe(resumen_tipo_formateado, use_container_width=True)

        # =========================================
        # RESUMEN POR SEDE
        # =========================================
        st.subheader("Resumen por sede")
        resumen_sede = df_filtrado.groupby("sede", dropna=False).agg(
            Cantidad=("sede", "size"),
            Costo_servicio=("Costo_servicio", "sum"),
            sobrecosto_total_espera=("sobrecosto_total_espera", "sum"),
            penalidad_origen=("penalidad_origen", "sum"),
            penalidad_destino=("penalidad_destino", "sum"),
            perdida_cita=("perdida_cita_monto", "sum"),
            penalidad_total=("penalidad_total", "sum")
        ).reset_index()

        fila_total_sede = {
            "sede": "TOTAL",
            "Cantidad": resumen_sede["Cantidad"].sum(),
            "Costo_servicio": resumen_sede["Costo_servicio"].sum(),
            "sobrecosto_total_espera": resumen_sede["sobrecosto_total_espera"].sum(),
            "penalidad_origen": resumen_sede["penalidad_origen"].sum(),
            "penalidad_destino": resumen_sede["penalidad_destino"].sum(),
            "perdida_cita": resumen_sede["perdida_cita"].sum(),
            "penalidad_total": resumen_sede["penalidad_total"].sum()
        }

        resumen_sede = pd.concat(
            [resumen_sede, pd.DataFrame([fila_total_sede])],
            ignore_index=True
        )

        resumen_sede_formateado = resumen_sede.copy()
        for col in ["Cantidad"]:
            resumen_sede_formateado[col] = resumen_sede_formateado[col].apply(
                lambda x: f"{x:,.0f}" if pd.notna(x) else ""
            )
        for col in ["Costo_servicio", "sobrecosto_total_espera", "penalidad_origen", "penalidad_destino", "perdida_cita", "penalidad_total"]:
            resumen_sede_formateado[col] = resumen_sede_formateado[col].apply(
                lambda x: f"S/ {x:,.2f}" if pd.notna(x) else ""
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
        tabla_cruzada_formateada = tabla_cruzada.copy()

        for col in tabla_cruzada_formateada.columns:
            if col == "Etiquetas de fila":
                continue
            if isinstance(col, tuple):
                if col[1] == "Cantidad":
                    tabla_cruzada_formateada[col] = tabla_cruzada_formateada[col].apply(lambda x: f"{x:,.0f}")
                elif col[1] == "Tarifa Base.":
                    tabla_cruzada_formateada[col] = tabla_cruzada_formateada[col].apply(lambda x: f"S/ {x:,.2f}")

        styler_cruzada = tabla_cruzada_formateada.style
        styler_cruzada = styler_cruzada.set_table_styles([
            {"selector": "th", "props": [("text-align", "center")]}
        ])

        primera_col = tabla_cruzada_formateada.columns[0]
        styler_cruzada = styler_cruzada.set_properties(
            subset=[primera_col],
            **{"text-align": "left"}
        )

        for col in tabla_cruzada_formateada.columns[1:]:
            styler_cruzada = styler_cruzada.set_properties(
                subset=[col],
                **{"text-align": "center"}
            )

        st.dataframe(styler_cruzada, use_container_width=True)

        # =========================================
        # DESCARGA
        # =========================================
        excel_bytes = exportar_excel(df_resultado)
        st.download_button(
            "Descargar resultado",
            data=excel_bytes,
            file_name="resultado_tiempo_espera_penalidades.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
