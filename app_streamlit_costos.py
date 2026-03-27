import math
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analisis Operativo de Servicios de Ambulancia", layout="wide")

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

TARIFA_PENALIDAD = {
    "TIPO I": 32.33,
    "TIPO II": 41.67,
    "TIPO III": 460.00,
    "TIPO III NEONATAL": 920.00,
}

# =========================================================
# CENTROS ASISTENCIALES PERMITIDOS
# =========================================================
CENTROS_ASISTENCIALES_PERMITIDOS = [
    "Centro de Atención Integral en Diabetes e Hipertensión",
    "Centro de Prevención de Riesgo del Trabajo (CEPRIT)",
    "Centro de Rehabilitación Profesional y Social (CERPS)",
    "Centro Nacional de Salud Renal (CNSR)",
    "Hospital Alberto L. Barton Thompson",
    "Hospital de Emergencias Grau",
    "Hospital Domingo Angulo",
    "Hospital Edgardo Rebagliati Martins",
    "Hospital Guillermo Almenara Irigoyen",
    "Hospital I Aurelio Díaz Ufano y Peral",
    "Hospital I Carlos Alcántara Butterfield",
    "Hospital II Ramón Castilla",
    "Hospital III Emergencias Grau",
    "Hospital III Suárez Angamos",
    "Hospital Nacional Alberto Sabogal Sologuren",
    "Hospital Octavio Mongrut Muñoz",
    "Hospital Uldarico Rocca Fernández",
    "Instituto Nacional Cardiovascular (INCOR)",
    "Policlínico Chincha",
    "Policlínico Francisco Pizarro",
    "Policlínico Juan José Rodríguez Lazo",
    "Policlínico Pablo Bermúdez",
    "Policlínico Próceres",
]

# =========================================================
# UTILIDADES
# =========================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    valor = valor.replace("–", "-").replace("—", "-")
    valor = " ".join(valor.split())
    return valor


def safe_round(valor, dec=2):
    if pd.isna(valor):
        return np.nan
    return round(float(valor), dec)


def parsear_fecha_segura(valor):
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, pd.Timestamp):
        return valor

    # Serial de Excel
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            return pd.to_datetime(valor, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            return pd.NaT

    txt = str(valor).strip()
    if not txt:
        return pd.NaT

    # Primero intenta formatos año-mes-día / año/mes/día
    formatos_ymd = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
    ]
    for fmt in formatos_ymd:
        try:
            return pd.to_datetime(txt, format=fmt, errors="raise")
        except Exception:
            pass

    # Luego intenta día/mes/año
    formatos_dmy = [
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
    ]
    for fmt in formatos_dmy:
        try:
            return pd.to_datetime(txt, format=fmt, errors="raise")
        except Exception:
            pass

    # Último intento automático
    try:
        return pd.to_datetime(txt, errors="coerce")
    except Exception:
        return pd.NaT


def parsear_hora_segura(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        return {
            "hour": valor.hour,
            "minute": valor.minute,
            "second": valor.second
        }

    # Serial Excel que representa hora
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            hora_dt = pd.to_datetime(valor, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notna(hora_dt):
                return {
                    "hour": hora_dt.hour,
                    "minute": hora_dt.minute,
                    "second": hora_dt.second
                }
        except Exception:
            pass

    txt = str(valor).strip()
    if not txt:
        return None

    try:
        partes = txt.split(":")
        if len(partes) >= 2:
            hh = int(partes[0])
            mm = int(partes[1])
            ss = int(partes[2]) if len(partes) > 2 else 0
            return {"hour": hh, "minute": mm, "second": ss}
    except Exception:
        pass

    try:
        hora_dt = pd.to_datetime(txt, errors="coerce")
        if pd.notna(hora_dt):
            return {
                "hour": hora_dt.hour,
                "minute": hora_dt.minute,
                "second": hora_dt.second
            }
    except Exception:
        pass

    return None


def parsear_columna_fecha(serie):
    return serie.apply(parsear_fecha_segura)


def combinar_fecha_hora(fecha_col, hora_col):
    if pd.isna(fecha_col) or pd.isna(hora_col):
        return pd.NaT

    fecha = parsear_fecha_segura(fecha_col)
    if pd.isna(fecha):
        return pd.NaT

    hora = parsear_hora_segura(hora_col)
    if hora is None:
        return pd.NaT

    return pd.Timestamp(
        year=fecha.year,
        month=fecha.month,
        day=fecha.day,
        hour=hora["hour"],
        minute=hora["minute"],
        second=hora["second"],
    )


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

def obtener_tarifa_penalidad(tipo_unidad):
    return TARIFA_PENALIDAD.get(normalizar_texto(tipo_unidad), 0.0)


def calcular_excedente_espera(minutos_espera):
    if pd.isna(minutos_espera) or minutos_espera <= MINUTOS_LIBRES_ESPERA:
        return 0.0
    return float(minutos_espera - MINUTOS_LIBRES_ESPERA)


def calcular_ocurrencias_espera(minutos_espera):
    excedente = calcular_excedente_espera(minutos_espera)
    if excedente <= 0:
        return 0
    return int(math.ceil(excedente / BLOQUE_ESPERA))


def agregar_fila_total(df_resumen, col_texto):
    if df_resumen.empty:
        return df_resumen.copy()

    df_total = df_resumen.copy()
    totales = df_total.select_dtypes(include="number").sum(numeric_only=True)
    fila_total = {col_texto: "TOTAL"}
    fila_total.update(totales.to_dict())
    return pd.concat([df_total, pd.DataFrame([fila_total])], ignore_index=True)


def formatear_resumen(df_resumen):
    df_fmt = df_resumen.copy()
    for col in df_fmt.select_dtypes(include="number").columns:
        df_fmt[col] = df_fmt[col].apply(lambda x: f"{x:,.2f}")
    return df_fmt


# =========================================================
# REGLAS DE TIEMPO DE ESPERA
# =========================================================
def calcular_tiempo_espera_origen(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")

    contacto_origen = row.get("contacto_paciente_origen")
    partida_origen = row.get("partida_origen")
    llegada_origen = row.get("llegada_origen")
    dt_programacion = row.get("dt_programacion")

    # 1) CITA / IDA / PROGRAMADA
    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

    # 2) CITA / IDA / NO PROGRAMADA
    if motivo == "CITA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

    # 3) CITA / RETORNO / PROGRAMADA o NO PROGRAMADA
    if motivo == "CITA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA", "AMBAS(POR ERROR)"]:
        return minutos_diff(dt_programacion, partida_origen)

    # 4) REFERENCIA / IDA / PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

    # 5) REFERENCIA / IDA / NO PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

    # 6) REFERENCIA / RETORNO / PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            if llegada_origen <= dt_programacion:
                return minutos_diff(dt_programacion, partida_origen)
            return minutos_diff(llegada_origen, partida_origen)
        return np.nan

    # 7) REFERENCIA / RETORNO / NO PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            if llegada_origen <= dt_programacion:
                return minutos_diff(dt_programacion, partida_origen)
            return minutos_diff(llegada_origen, partida_origen)
        return np.nan

    # 8) EMERGENCIA / IDA / NO PROGRAMADA
    if motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

    # 9) ALTA / IDA / NO PROGRAMADA
    if motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(contacto_origen, partida_origen)

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

    # CITA / RETORNO
    if motivo == "CITA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA", "AMBAS(POR ERROR)"]:
        if pd.notna(contacto_origen) and pd.notna(dt_programacion):
            if contacto_origen > dt_programacion:
                return 0.0

    # REFERENCIA / RETORNO / PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            if llegada_origen > dt_programacion:
                return 0.0

    # REFERENCIA / RETORNO / NO PROGRAMADA
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

    # 1) CITA / IDA / PROGRAMADA
    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            if llegada_destino <= dt_programacion:
                return minutos_diff(dt_programacion, hora_finalizacion)
            return minutos_diff(llegada_destino, hora_finalizacion)
        return np.nan

    # 2) CITA / IDA / NO PROGRAMADA
    if motivo == "CITA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(dt_programacion, hora_finalizacion)

    # 3) CITA / RETORNO / PROGRAMADA o NO PROGRAMADA
    if motivo == "CITA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA", "AMBAS(POR ERROR)"]:
        return minutos_diff(llegada_destino, hora_finalizacion)

    # 4) REFERENCIA / IDA / PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            if llegada_destino <= dt_programacion:
                return minutos_diff(dt_programacion, hora_finalizacion)
            return minutos_diff(llegada_destino, hora_finalizacion)
        return np.nan

    # 5) REFERENCIA / IDA / NO PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    # 6) REFERENCIA / RETORNO / PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    # 7) REFERENCIA / RETORNO / NO PROGRAMADA
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    # 8) EMERGENCIA / IDA / NO PROGRAMADA
    if motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return minutos_diff(llegada_destino, hora_finalizacion)

    # 9) ALTA / IDA / NO PROGRAMADA
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

    # CITA / IDA / PROGRAMADA
    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            dif = minutos_diff(dt_programacion, llegada_origen)
            if pd.notna(dif) and dif < 60:
                return 0.0

    return minutos

# =========================
#  PENALIDADES
# =========================
def calcular_penalidades(row):
    motivo = row.get("motivo_traslado", "")
    sentido = row.get("sentido_traslado", "")
    modalidad = row.get("modalidad", "")
    tipo_unidad = row.get("tipo_unidad", "")

    llegada_origen = row.get("llegada_origen")
    contacto = row.get("contacto_paciente_origen")
    llegada_destino = row.get("llegada_destino")
    dt_programacion = row.get("dt_programacion")
    dt_registro = row.get("dt_registro")

    penalidad_origen = 0.0
    penalidad_destino = 0.0
    detalle = ""

    tarifa_penalidad = obtener_tarifa_penalidad(tipo_unidad)

    def bloques(minutos, tam_bloque=30):
        if pd.isna(minutos) or minutos <= 0:
            return 0
        return int(math.ceil(minutos / tam_bloque))

    # =========================
    # CITA RETORNO (PROGRAMADA / NO PROGRAMADA)
    # =========================
    if motivo == "CITA" and sentido == "RETORNO":
        if pd.notna(contacto) and pd.notna(dt_programacion):
            atraso = minutos_diff(dt_programacion, contacto)  # contacto - programacion
            if pd.notna(atraso) and atraso > 0:
                b = bloques(atraso, 30)
                penalidad_origen = b * tarifa_penalidad
                detalle = f"CITA RETORNO atraso {safe_round(atraso)} min"

    # =========================
    # REFERENCIA IDA PROGRAMADA (DESTINO)
    # =========================
    elif motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        if pd.notna(llegada_destino) and pd.notna(dt_programacion):
            atraso = minutos_diff(dt_programacion, llegada_destino)  # llegada - programacion
            if pd.notna(atraso) and atraso > 0:
                llegada_origen_ref = row.get("llegada_origen")
                if pd.notna(llegada_origen_ref):
                    anticipo = minutos_diff(llegada_origen_ref, dt_programacion)  # programacion - llegada_origen
                    if pd.notna(anticipo) and anticipo >= 90:
                        penalidad_destino = 0.0
                        detalle = "REFERENCIA IDA PROG exonerada por anticipación >= 90 min"
                    else:
                        b = bloques(atraso, 30)
                        penalidad_destino = b * tarifa_penalidad
                        detalle = f"REF IDA PROG atraso {safe_round(atraso)} min"
                else:
                    b = bloques(atraso, 30)
                    penalidad_destino = b * tarifa_penalidad
                    detalle = f"REF IDA PROG atraso {safe_round(atraso)} min"

    # =========================
    # REFERENCIA IDA NO PROGRAMADA (ORIGEN)
    # =========================
    elif motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_registro):

            # No penaliza si es tipo III
            if normalizar_texto(tipo_unidad) == "TIPO III":
                penalidad_origen = 0.0
                detalle = "REF IDA NO PROG - SIN PENALIDAD (TIPO III)"
            else:
                minutos = minutos_diff(dt_registro, llegada_origen)
                if pd.notna(minutos) and minutos > 30:
                    exceso = minutos - 30
                    b = bloques(exceso, 30)
                    penalidad_origen = b * tarifa_penalidad
                    detalle = f"REF IDA NO PROG exceso {safe_round(exceso)} min"

    # =========================
    # REFERENCIA RETORNO PROGRAMADA
    # =========================
    elif motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        if pd.notna(contacto) and pd.notna(dt_programacion):
            atraso = minutos_diff(dt_programacion, contacto)
            if pd.notna(atraso) and atraso > 0:
                gracia = 30 if dt_programacion.year >= 2026 else 0
                exceso = max(0, atraso - gracia)
                if exceso > 0:
                    b = bloques(exceso, 30)
                    penalidad_origen = b * tarifa_penalidad
                    detalle = f"REF RET PROG exceso {safe_round(exceso)} min (gracia {gracia})"

    # =========================
    # REFERENCIA RETORNO NO PROGRAMADA
    # =========================
    elif motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            atraso = minutos_diff(dt_programacion, llegada_origen)
            if pd.notna(atraso) and atraso > 0:
                gracia = 30 if dt_programacion.year >= 2026 else 0
                exceso = max(0, atraso - gracia)
                if exceso > 0:
                    b = bloques(exceso, 30)
                    penalidad_origen = b * tarifa_penalidad
                    detalle = f"REF RET NO PROG exceso {safe_round(exceso)} min (gracia {gracia})"

    # =========================
    # EMERGENCIA IDA NO PROGRAMADA
    # =========================
    elif motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_registro):
            minutos = minutos_diff(dt_registro, llegada_origen)

            sede = normalizar_texto(row.get("sede", ""))
            tolerancia = 15 if "BARTON" in sede else 30
            bloque = 15 if "BARTON" in sede else 30

            if pd.notna(minutos) and minutos > tolerancia:
                exceso = minutos - tolerancia
                b = bloques(exceso, bloque)
                penalidad_origen = b * tarifa_penalidad
                detalle = f"EMERGENCIA exceso {safe_round(exceso)} min"

    # =========================
    # ALTA IDA NO PROGRAMADA
    # =========================
    elif motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        if pd.notna(llegada_origen) and pd.notna(dt_registro):
            minutos = minutos_diff(dt_registro, llegada_origen)
            if pd.notna(minutos) and minutos > 30:
                exceso = minutos - 30
                b = bloques(exceso, 30)
                penalidad_origen = b * tarifa_penalidad
                detalle = f"ALTA exceso {safe_round(exceso)} min"

    return pd.Series({
        "penalidad_origen": round(penalidad_origen, 2),
        "penalidad_destino": round(penalidad_destino, 2),
        "penalidad_total": round(penalidad_origen + penalidad_destino, 2),
        "detalle_penalidad": detalle
    })

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
def procesar_archivo(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).lower().replace(" ", "_") for c in df.columns]

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
            "NO PROGRAMADOS": "NO PROGRAMADA",
            "NO PROGRAMADA ": "NO PROGRAMADA",
            "AMBAS(POR ERROR)": "AMBAS(POR ERROR)"
        })

    columnas_datetime = [
        "salida_de_base",
        "llegada_origen",
        "contacto_paciente_origen",
        "partida_origen",
        "llegada_destino",
        "contacto_paciente_destino",
        "hora_finalizacion",
        "fecha_registro",
        "fecha_programada",
    ]

    for col in columnas_datetime:
        if col in df.columns:
            df[col] = parsear_columna_fecha(df[col])

    if "dt_registro" not in df.columns:
        df["dt_registro"] = pd.NaT

    df["dt_registro"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_registro"), r.get("hora_registro")),
        axis=1
    )
    df["dt_programacion"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_programada"), r.get("hora_programada")),
        axis=1
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

   # =========================
   # APLICAR PENALIDADES
   # =========================
    resultado_penalidad = df_salida.apply(calcular_penalidades, axis=1)
    df_salida = pd.concat([df_salida, resultado_penalidad], axis=1)

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
st.title("Analisis Operativo de Servicios de Ambulancia")
st.caption("Análisis de costos, tiempos de espera y penalidades")

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
        # =========================
        # FILTROS TIPO DASHBOARD
        # =========================
        st.subheader("Filtros")

        f1, f2, f3, f4 = st.columns(4)

        with f1:
            estados = sorted(df_resultado["estado"].dropna().unique().tolist())
            estado_sel = st.multiselect(
                "Estado",
                options=estados,
                default=[],
                placeholder="Todos"
            )

        with f2:
            sedes = sorted(df_resultado["sede"].dropna().unique().tolist())
            sede_sel = st.multiselect(
                "Sede",
                options=sedes,
                default=[],
                placeholder="Todos"
            )

        with f3:
            motivos = sorted(df_resultado["motivo_traslado"].dropna().unique().tolist())
            motivo_sel = st.multiselect(
                "Motivo",
                options=motivos,
                default=[],
                placeholder="Todos"
            )

        with f4:
            tipos = sorted(df_resultado["tipo_unidad"].dropna().unique().tolist())
            tipo_sel = st.multiselect(
                "Tipo unidad",
                options=tipos,
                default=[],
                placeholder="Todos"
            )

        b1, b2 = st.columns([2, 1])

        with b1:
            buscar = st.text_input("Buscar Nro Solicitud")

        with b2:
            solo_penalidad = st.checkbox("Solo con penalidad")

        # =========================
        # APLICAR FILTROS (MULTISELECT)
        # =========================
        df_filtrado = df_resultado.copy()

        if "estado" in df_filtrado.columns and estado_sel:
            df_filtrado = df_filtrado[df_filtrado["estado"].isin(estado_sel)]

        if "sede" in df_filtrado.columns and sede_sel:
            df_filtrado = df_filtrado[df_filtrado["sede"].isin(sede_sel)]

        if "motivo_traslado" in df_filtrado.columns and motivo_sel:
            df_filtrado = df_filtrado[df_filtrado["motivo_traslado"].isin(motivo_sel)]

        if "tipo_unidad" in df_filtrado.columns and tipo_sel:
            df_filtrado = df_filtrado[df_filtrado["tipo_unidad"].isin(tipo_sel)]

        if buscar and "nro_solicitud" in df_filtrado.columns:
            df_filtrado = df_filtrado[
                df_filtrado["nro_solicitud"].astype(str).str.contains(buscar, case=False, na=False)
            ]

        if solo_penalidad and "penalidad_total" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["penalidad_total"] > 0]

        st.success("Archivo procesado correctamente.")
        # =========================
        # KPIs
        # =========================
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Registros", len(df_filtrado))
        c2.metric("Costo servicio total", f"{df_filtrado['Costo_servicio'].sum():,.2f}")
        c3.metric("Sobrecosto total espera", f"{df_filtrado['sobrecosto_total_espera'].sum():,.2f}")
        c4.metric("Penalidad total", f"{df_filtrado['penalidad_total'].sum():,.2f}")

        # =========================
        # TOP CASOS
        # =========================
        st.subheader("Top casos")

        t1, t2 = st.columns(2)

        with t1:
            top_espera = df_filtrado.sort_values("sobrecosto_total_espera", ascending=False).head(10)
            cols_top_espera = [c for c in [
                "nro_solicitud", "tipo_unidad", "motivo_traslado", "sede",
                "sobrecosto_total_espera", "tiempo_espera_total"
            ] if c in top_espera.columns]
            st.caption("Top 10 por sobrecosto de espera")
            st.dataframe(top_espera[cols_top_espera], use_container_width=True)

        with t2:
            top_penalidad = df_filtrado.sort_values("penalidad_total", ascending=False).head(10)
            cols_top_pen = [c for c in [
                "nro_solicitud", "tipo_unidad", "motivo_traslado", "sede",
                "penalidad_total", "detalle_penalidad"
            ] if c in top_penalidad.columns]
            st.caption("Top 10 por penalidad")
            st.dataframe(top_penalidad[cols_top_pen], use_container_width=True)

        # =========================
        # DETALLE
        # =========================
        st.subheader("Vista previa")
        st.dataframe(df_filtrado.head(20), use_container_width=True)

        # =========================
        # RESUMEN POR TIPO
        # =========================
        st.subheader("Resumen por tipo de unidad")
        resumen_tipo = df_filtrado.groupby("tipo_unidad", dropna=False).agg(
            {
                "Costo_servicio": "sum",
                "sobrecosto_total_espera": "sum",
                "penalidad_total": "sum",
            }
        ).reset_index()
        resumen_tipo = agregar_fila_total(resumen_tipo, "tipo_unidad")
        st.dataframe(formatear_resumen(resumen_tipo), use_container_width=True)

        # =========================
        # RESUMEN POR SEDE
        # =========================
        st.subheader("Resumen por sede")
        resumen_sede = df_filtrado.groupby("sede", dropna=False).agg(
            {
                "Costo_servicio": "sum",
                "sobrecosto_total_espera": "sum",
                "penalidad_total": "sum",
            }
        ).reset_index()
        resumen_sede = agregar_fila_total(resumen_sede, "sede")
        st.dataframe(formatear_resumen(resumen_sede), use_container_width=True)
# =========================
# RESUMEN POR C_ASISTENCIAL_DESTINO
# =========================
st.subheader("Resumen por centro asistencial de destino")

if "c_asistencial_destino" in df_filtrado.columns:
    centros_permitidos_set = {
        normalizar_texto(nombre) for nombre in CENTROS_ASISTENCIALES_PERMITIDOS
    }

    resumen_destino = (
        df_filtrado.groupby("c_asistencial_destino", dropna=False)
        .agg(
            ocurrencias=("c_asistencial_destino", "count"),
            tiempo_excedente_total=("minutos_excedentes_destino", "sum"),
            tiempo_espera_total=("tiempo_espera_total", "sum"),
        )
        .reset_index()
    )

    resumen_destino["c_asistencial_destino"] = resumen_destino["c_asistencial_destino"].fillna("SIN DATO")
    resumen_destino["c_asistencial_destino_normalizado"] = resumen_destino["c_asistencial_destino"].apply(normalizar_texto)

    resumen_destino = resumen_destino[
        resumen_destino["c_asistencial_destino_normalizado"].isin(centros_permitidos_set)
    ].copy()

    orden_centros = {
        normalizar_texto(nombre): idx
        for idx, nombre in enumerate(CENTROS_ASISTENCIALES_PERMITIDOS)
    }

    resumen_destino["orden"] = resumen_destino["c_asistencial_destino_normalizado"].map(orden_centros)
    resumen_destino = resumen_destino.sort_values("orden").drop(
        columns=["c_asistencial_destino_normalizado", "orden"]
    )

    resumen_destino = agregar_fila_total(resumen_destino, "c_asistencial_destino")
    st.dataframe(formatear_resumen(resumen_destino), use_container_width=True)
     

        # =========================
        # DESCARGA
        # =========================
        excel_bytes = exportar_excel(df_filtrado)
        st.download_button(
            "Descargar resultado filtrado",
            data=excel_bytes,
            file_name="resultado_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Archivo procesado correctamente.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
