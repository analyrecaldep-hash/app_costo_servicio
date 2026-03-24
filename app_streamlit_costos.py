import math
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Cálculo web de costo por servicio", layout="wide")

# =========================================================
# CONFIGURACIÓN
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

TARIFAS_PENALIDAD = {
    "TIPO I": 32.33,
    "TIPO II": 41.67,
    "TIPO III": 460.00,
    "TIPO III NEONATAL": 920.00,
}

MONTO_PERDIDA_CITA = 535.00

MAX_MINUTOS_PENALIZABLES = 12 * 60
MAX_MINUTOS_DIFERENCIA = 3 * 24 * 60


# =========================================================
# UTILIDADES
# =========================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    return " ".join(valor.split())


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


def obtener_tarifa_penalidad(tipo_unidad):
    return TARIFAS_PENALIDAD.get(normalizar_texto(tipo_unidad), 0.0)


def calcular_excedente_espera(minutos_espera):
    if pd.isna(minutos_espera) or minutos_espera <= MINUTOS_LIBRES_ESPERA:
        return 0.0
    return float(minutos_espera - MINUTOS_LIBRES_ESPERA)


def calcular_ocurrencias_espera(minutos_espera):
    excedente = calcular_excedente_espera(minutos_espera)
    if excedente <= 0:
        return 0
    return int(math.ceil(excedente / BLOQUE_ESPERA))


def calcular_bloques(minutos, tam_bloque=30):
    if pd.isna(minutos) or minutos <= 0:
        return 0
    return int(math.ceil(minutos / float(tam_bloque)))


def es_policlinico_barton(row):
    candidatos = [
        row.get("origen", ""),
        row.get("lugar_origen", ""),
        row.get("establecimiento_origen", ""),
        row.get("sede", ""),
    ]
    texto = " ".join([normalizar_texto(x) for x in candidatos])
    return "POLICLINICO BARTON" in texto or ("BARTON" in texto and "POLICLINICO" in texto)


def obtener_dt_registro(row):
    candidatos_directos = ["dt_registro", "fecha_hora_registro", "registro"]
    for col in candidatos_directos:
        if col in row.index and pd.notna(row.get(col)):
            return row.get(col)

    fecha_reg = row.get("fecha_registro", pd.NaT)
    hora_reg = row.get("hora_registro", pd.NaT)
    if pd.notna(fecha_reg) and pd.notna(hora_reg):
        return combinar_fecha_hora(fecha_reg, hora_reg)

    return pd.NaT


def es_diferencia_inconsistente(minutos):
    if pd.isna(minutos):
        return False
    return abs(minutos) > MAX_MINUTOS_DIFERENCIA


def cap_minutos_penalizables(minutos):
    if pd.isna(minutos) or minutos <= 0:
        return 0.0
    return min(float(minutos), float(MAX_MINUTOS_PENALIZABLES))


# =========================================================
# NORMALIZACIÓN
# =========================================================
def normalizar_dataframe(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).lower().replace(" ", "_") for c in df.columns]

    columnas_texto = [
        "nro_solicitud",
        "sentido_traslado", "sede", "motivo_traslado", "modalidad", "estado",
        "efectivo", "tipo_unidad", "origen", "lugar_origen", "establecimiento_origen"
    ]

    for col in columnas_texto:
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
        "fecha_programada",
        "fecha_registro",
        "dt_registro",
        "fecha_hora_registro",
        "registro",
    ]

    for col in columnas_datetime:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    if "hora_programada" not in df.columns:
        df["hora_programada"] = pd.NaT
    if "fecha_programada" not in df.columns:
        df["fecha_programada"] = pd.NaT

    if "dt_registro" not in df.columns:
        df["dt_registro"] = pd.NaT

    df["dt_programacion"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_programada"), r.get("hora_programada")),
        axis=1
    )

    df["dt_registro_calc"] = df.apply(obtener_dt_registro, axis=1)

    mask_dt_registro_vacio = df["dt_registro_calc"].isna()
    if mask_dt_registro_vacio.any():
        df.loc[mask_dt_registro_vacio, "dt_registro_calc"] = df.loc[mask_dt_registro_vacio].apply(
            lambda r: combinar_fecha_hora(r.get("fecha_registro"), r.get("hora_registro")),
            axis=1
        )

    return df


# =========================================================
# LÓGICA DE ESPERA
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
    dt_registro = row.get("dt_registro_calc")

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
# CLASIFICACIÓN DE REGLAS DE PENALIDAD
# =========================================================
def obtener_codigo_regla(row):
    motivo = normalizar_texto(row.get("motivo_traslado"))
    sentido = normalizar_texto(row.get("sentido_traslado"))
    modalidad = normalizar_texto(row.get("modalidad"))

    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return "CITA_IDA_PROGRAMADA"
    if motivo == "CITA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "CITA_IDA_NO_PROGRAMADA"
    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        return "CITA_RETORNO_PROGRAMADA"
    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        return "CITA_RETORNO_NO_PROGRAMADA"
    if motivo == "CITA" and sentido == "RETORNO" and modalidad == "AMBAS(POR ERROR)":
        return "CITA_RETORNO_NO_PROGRAMADA"
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return "REFERENCIA_IDA_PROGRAMADA"
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "REFERENCIA_IDA_NO_PROGRAMADA"
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        return "REFERENCIA_RETORNO_PROGRAMADA"
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        return "REFERENCIA_RETORNO_NO_PROGRAMADA"
    if motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "EMERGENCIA_IDA_NO_PROGRAMADA"
    if motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "ALTA_IDA_NO_PROGRAMADA"

    return "SIN_REGLA"


def fila_resultado_base(codigo_regla):
    return {
        "codigo_regla": codigo_regla,
        "regla_aplicada": "",
        "observacion_calculo": "",
        "diferencia_penalidad_origen": np.nan,
        "diferencia_penalidad_destino": np.nan,
        "min_penalidad_origen": 0.0,
        "min_penalidad_destino": 0.0,
        "bloques_penalidad_origen": 0,
        "bloques_penalidad_destino": 0,
        "penalidad_origen": 0.0,
        "penalidad_destino": 0.0,
        "perdida_cita_flag": 0,
        "perdida_cita_monto": 0.0,
        "flag_inconsistencia_fecha": 0,
    }


# =========================================================
# FUNCIONES ESPECÍFICAS DE PENALIDAD
# =========================================================
def calcular_penalidad_cita_retorno(row, res):
    dt_programacion = row.get("dt_programacion")
    contacto_origen = row.get("contacto_paciente_origen")
    costo_servicio = pd.to_numeric(row.get("Costo_servicio", np.nan), errors="coerce")

    res["regla_aplicada"] = "CITA RETORNO"

    if pd.isna(dt_programacion) or pd.isna(contacto_origen):
        res["observacion_calculo"] = "No se pudo calcular: falta hora de programación o contacto con paciente."
        return pd.Series(res)

    if pd.isna(costo_servicio):
        res["observacion_calculo"] = "No se pudo calcular: costo_servicio inválido."
        return pd.Series(res)

    retraso_real = minutos_diff(dt_programacion, contacto_origen)
    res["diferencia_penalidad_origen"] = safe_round(retraso_real)

    if es_diferencia_inconsistente(retraso_real):
        res["flag_inconsistencia_fecha"] = 1
        res["observacion_calculo"] = "Diferencia inconsistente entre contacto_paciente_origen y programación."
        return pd.Series(res)

    if contacto_origen <= dt_programacion:
        res["min_penalidad_origen"] = 0.0
        res["bloques_penalidad_origen"] = 0
        res["penalidad_origen"] = 0.0
        res["observacion_calculo"] = (
            "Sin penalidad: hora de contacto con paciente dentro o antes del horario de programación."
        )
        return pd.Series(res)

    retraso_penalizable = cap_minutos_penalizables(retraso_real)
    bloques = calcular_bloques(retraso_penalizable, 30)
    penalidad_por_bloque = costo_servicio / 3.0

    res["min_penalidad_origen"] = safe_round(retraso_penalizable)
    res["bloques_penalidad_origen"] = bloques
    res["penalidad_origen"] = safe_round(bloques * penalidad_por_bloque)
    res["observacion_calculo"] = (
        "Se cobra 1/3 del costo del servicio por cada 30 minutos o fracción "
        "de retraso respecto a la hora de programación."
    )
    return pd.Series(res)


# =========================================================
# REGLAS DE PENALIDAD
# =========================================================
def calcular_penalidades_fila(row):
    codigo = obtener_codigo_regla(row)
    tarifa = obtener_tarifa_penalidad(row.get("tipo_unidad"))
    res = fila_resultado_base(codigo)

    dt_programacion = row.get("dt_programacion")
    dt_registro = row.get("dt_registro_calc")
    llegada_origen = row.get("llegada_origen")
    llegada_destino = row.get("llegada_destino")
    contacto_origen = row.get("contacto_paciente_origen")
    estado = row.get("estado", "")

    if codigo == "CITA_IDA_PROGRAMADA":
        res["regla_aplicada"] = "CITA IDA PROGRAMADA"

        if pd.notna(dt_programacion) and pd.notna(llegada_destino):
            hora_limite = dt_programacion - pd.Timedelta(minutes=15)
            atraso_real = minutos_diff(hora_limite, llegada_destino)
            res["diferencia_penalidad_destino"] = safe_round(atraso_real)

            if es_diferencia_inconsistente(atraso_real):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre hora límite y llegada_destino."
                return pd.Series(res)

            if pd.notna(atraso_real) and atraso_real > 0:
                atraso_penalizable = cap_minutos_penalizables(atraso_real)
                bloques = calcular_bloques(atraso_penalizable, 30)
                res["min_penalidad_destino"] = safe_round(atraso_penalizable)
                res["bloques_penalidad_destino"] = bloques
                res["penalidad_destino"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Se compara llegada_destino con (hora_programada - 15 min)."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó dentro del tiempo permitido."

        if pd.notna(dt_programacion) and pd.notna(llegada_origen):
            anticipacion = minutos_diff(llegada_origen, dt_programacion)

            if es_diferencia_inconsistente(anticipacion):
                res["flag_inconsistencia_fecha"] = 1
                res["min_penalidad_destino"] = 0.0
                res["bloques_penalidad_destino"] = 0
                res["penalidad_destino"] = 0.0
                res["observacion_calculo"] = "Diferencia inconsistente entre llegada_origen y programación."
                return pd.Series(res)

            if pd.notna(anticipacion) and anticipacion > 60:
                res["min_penalidad_destino"] = 0.0
                res["bloques_penalidad_destino"] = 0
                res["penalidad_destino"] = 0.0
                res["observacion_calculo"] = "Penalidad eliminada: programación - llegada_origen > 60 min."

        return pd.Series(res)

    if codigo == "CITA_IDA_NO_PROGRAMADA":
        res["regla_aplicada"] = "CITA IDA NO PROGRAMADA"
        res["observacion_calculo"] = "No aplica penalidad según reglas de negocio."
        return pd.Series(res)

    if codigo in ["CITA_RETORNO_PROGRAMADA", "CITA_RETORNO_NO_PROGRAMADA"]:
        return calcular_penalidad_cita_retorno(row, res)

    if codigo == "REFERENCIA_IDA_PROGRAMADA":
        res["regla_aplicada"] = "REFERENCIA IDA PROGRAMADA"

        if pd.isna(llegada_destino) and estado in ["VALIDADO", "CERRADO", "ATENDIDO", "FINALIZADO"]:
            res["perdida_cita_flag"] = 1
            res["perdida_cita_monto"] = MONTO_PERDIDA_CITA
            res["observacion_calculo"] = "Pérdida de cita: no existe llegada_destino."
            return pd.Series(res)

        if pd.notna(dt_programacion) and pd.notna(llegada_destino):
            atraso_real = minutos_diff(dt_programacion, llegada_destino)
            res["diferencia_penalidad_destino"] = safe_round(atraso_real)

            if es_diferencia_inconsistente(atraso_real):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre programación y llegada_destino."
                return pd.Series(res)

            if pd.notna(llegada_origen):
                anticipacion = minutos_diff(llegada_origen, dt_programacion)

                if es_diferencia_inconsistente(anticipacion):
                    res["flag_inconsistencia_fecha"] = 1
                    res["observacion_calculo"] = "Diferencia inconsistente entre llegada_origen y programación."
                    return pd.Series(res)

                if pd.notna(anticipacion) and anticipacion >= 90:
                    res["observacion_calculo"] = "Sin penalidad por exoneración: programación - llegada_origen >= 90 min."
                    return pd.Series(res)

            if pd.notna(atraso_real) and atraso_real > 0:
                atraso_penalizable = cap_minutos_penalizables(atraso_real)
                bloques = calcular_bloques(atraso_penalizable, 30)
                res["min_penalidad_destino"] = safe_round(atraso_penalizable)
                res["bloques_penalidad_destino"] = bloques
                res["penalidad_destino"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Se compara llegada_destino con hora_programada."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó dentro del horario."

        return pd.Series(res)

    if codigo == "REFERENCIA_IDA_NO_PROGRAMADA":
        res["regla_aplicada"] = "REFERENCIA IDA NO PROGRAMADA"

        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)

            if es_diferencia_inconsistente(demora_total):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre registro y llegada_origen."
                return pd.Series(res)

            if pd.notna(demora_total) and demora_total > 30:
                exceso = cap_minutos_penalizables(demora_total - 30)
                bloques = calcular_bloques(exceso, 30)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Máximo 30 min desde registro hasta llegada_origen."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó a origen dentro de los 30 min desde el registro."

        return pd.Series(res)

    if codigo == "REFERENCIA_RETORNO_PROGRAMADA":
        res["regla_aplicada"] = "REFERENCIA RETORNO PROGRAMADA"

        if pd.notna(contacto_origen) and pd.notna(dt_programacion):
            atraso_real = minutos_diff(dt_programacion, contacto_origen)
            res["diferencia_penalidad_origen"] = safe_round(atraso_real)

            if es_diferencia_inconsistente(atraso_real):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre contacto_paciente_origen y programación."
                return pd.Series(res)

            if pd.notna(atraso_real) and atraso_real > 0:
                gracia = 30 if dt_programacion.year >= 2026 else 0
                atraso_penalizable = max(0, atraso_real - gracia)
                atraso_penalizable = cap_minutos_penalizables(atraso_penalizable)

                if atraso_penalizable > 0:
                    bloques = calcular_bloques(atraso_penalizable, 30)
                    res["min_penalidad_origen"] = safe_round(atraso_penalizable)
                    res["bloques_penalidad_origen"] = bloques
                    res["penalidad_origen"] = safe_round(bloques * tarifa)
                    res["observacion_calculo"] = f"Se aplicó gracia de {gracia} min."
                else:
                    res["observacion_calculo"] = f"Sin penalidad: el atraso quedó cubierto por la gracia de {gracia} min."
            else:
                res["observacion_calculo"] = "Sin penalidad: contacto dentro del horario."

        return pd.Series(res)

    if codigo == "EMERGENCIA_IDA_NO_PROGRAMADA":
        res["regla_aplicada"] = "EMERGENCIA IDA NO PROGRAMADA"

        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)

            if es_diferencia_inconsistente(demora_total):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre registro y llegada_origen."
                return pd.Series(res)

            barton = es_policlinico_barton(row)
            tolerancia = 15 if barton else 30
            bloque = 15 if barton else 30

            if pd.notna(demora_total) and demora_total > tolerancia:
                exceso = cap_minutos_penalizables(demora_total - tolerancia)
                bloques = calcular_bloques(exceso, bloque)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Penalidad por llegada a origen fuera del tiempo máximo."
            else:
                res["observacion_calculo"] = "Sin penalidad."

        return pd.Series(res)

    if codigo == "REFERENCIA_RETORNO_NO_PROGRAMADA":
        res["regla_aplicada"] = "REFERENCIA RETORNO NO PROGRAMADA"

        if pd.notna(llegada_origen) and pd.notna(dt_registro):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)

            if es_diferencia_inconsistente(demora_total):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre registro y llegada_origen."
                return pd.Series(res)

            if pd.notna(demora_total) and demora_total > 30:
                exceso = cap_minutos_penalizables(demora_total - 30)
                bloques = calcular_bloques(exceso, 30)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Máximo 30 min desde registro hasta llegada_origen."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó a origen dentro de los 30 min desde el registro."

        return pd.Series(res)

    if codigo == "ALTA_IDA_NO_PROGRAMADA":
        res["regla_aplicada"] = "ALTA IDA NO PROGRAMADA"

        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)

            if es_diferencia_inconsistente(demora_total):
                res["flag_inconsistencia_fecha"] = 1
                res["observacion_calculo"] = "Diferencia inconsistente entre registro y llegada_origen."
                return pd.Series(res)

            if pd.notna(demora_total) and demora_total > 30:
                exceso = cap_minutos_penalizables(demora_total - 30)
                bloques = calcular_bloques(exceso, 30)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["observacion_calculo"] = "Máximo 30 min desde registro hasta llegada_origen."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó a origen dentro de los 30 min desde el registro."

        return pd.Series(res)

    res["observacion_calculo"] = "No existe regla configurada para esta combinación."
    return pd.Series(res)


def procesar_penalidades(df):
    df_pen = normalizar_dataframe(df).copy()

    resultado = df_pen.apply(calcular_penalidades_fila, axis=1).reset_index(drop=True)

    df_resultado = pd.concat(
        [df_pen[["nro_solicitud"]].reset_index(drop=True), resultado],
        axis=1
    )

    df_resultado["penalidad_total"] = (
        df_resultado["penalidad_origen"].fillna(0)
        + df_resultado["penalidad_destino"].fillna(0)
        + df_resultado["perdida_cita_monto"].fillna(0)
    ).round(2)

    return df_resultado


# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
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


def procesar_archivo(df):
    df = normalizar_dataframe(df)

    resultado_espera = df.apply(procesar_fila, axis=1)
    df_salida = pd.concat([df, resultado_espera], axis=1)

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

    resultado_penalidades = procesar_penalidades(df_salida)

    columnas_penalidad = [
        "nro_solicitud",
        "codigo_regla",
        "regla_aplicada",
        "observacion_calculo",
        "diferencia_penalidad_origen",
        "diferencia_penalidad_destino",
        "min_penalidad_origen",
        "min_penalidad_destino",
        "bloques_penalidad_origen",
        "bloques_penalidad_destino",
        "penalidad_origen",
        "penalidad_destino",
        "perdida_cita_flag",
        "perdida_cita_monto",
        "flag_inconsistencia_fecha",
        "penalidad_total",
    ]

    resultado_penalidades = resultado_penalidades[columnas_penalidad]

    df_final = df_salida.merge(
        resultado_penalidades,
        on="nro_solicitud",
        how="left",
        suffixes=("", "_pen")
    )

    return df_final


def convertir_a_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
    output.seek(0)
    return output


# =========================================================
# INTERFAZ STREAMLIT
# =========================================================
st.title("Cálculo de costo de servicio, tiempo de espera y penalidades")

archivo = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xls"]
)

if archivo is not None:
    try:
        df_original = pd.read_excel(archivo)
        st.subheader("Vista previa")
        st.dataframe(df_original.head(20), use_container_width=True)

        if st.button("Procesar archivo"):
            df_resultado = procesar_archivo(df_original)

            st.success("Archivo procesado correctamente")
            st.subheader("Resultado")
            st.dataframe(df_resultado, use_container_width=True)

            excel_bytes = convertir_a_excel(df_resultado)
            st.download_button(
                label="Descargar resultado en Excel",
                data=excel_bytes,
                file_name="resultado_calculo_penalidades.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
