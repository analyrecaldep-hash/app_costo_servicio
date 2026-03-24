
import math
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd

TARIFAS_PENALIDAD = {
    "TIPO I": 32.33,
    "TIPO II": 41.67,
    "TIPO III": 460.00,
    "TIPO III NEONATAL": 920.00,
}

MONTO_PERDIDA_CITA = 535.00


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


def obtener_tarifa(tipo_unidad):
    return TARIFAS_PENALIDAD.get(normalizar_texto(tipo_unidad), 0.0)


def calcular_bloques(minutos, tam_bloque):
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
        if col in row and pd.notna(row.get(col)):
            return row.get(col)

    fecha_reg = row.get("fecha_registro", pd.NaT)
    hora_reg = row.get("hora_registro", pd.NaT)
    if pd.notna(fecha_reg) and pd.notna(hora_reg):
        return combinar_fecha_hora(fecha_reg, hora_reg)

    return pd.NaT


def normalizar_datos(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).lower().replace(" ", "_") for c in df.columns]

    columnas_texto = [
        "motivo_traslado", "sentido_traslado", "modalidad", "estado",
        "tipo_unidad", "sede", "origen", "lugar_origen", "establecimiento_origen"
    ]
    for col in columnas_texto:
        if col in df.columns:
            df[col] = df[col].apply(normalizar_texto)

    if "motivo_traslado" in df.columns:
        df["motivo_traslado"] = df["motivo_traslado"].replace({
            "REFERENCIAS": "REFERENCIA",
            "EMERGENCIAS": "EMERGENCIA",
            "ALTAS": "ALTA",
        })

    if "modalidad" in df.columns:
        df["modalidad"] = df["modalidad"].replace({
            "PROGRAMADAS": "PROGRAMADA",
            "NO PROGRAMADO": "NO PROGRAMADA",
            "NO PROGRAMADOS": "NO PROGRAMADA",
            "NO PROGRAMADA ": "NO PROGRAMADA",
        })

    columnas_datetime = [
        "fecha_programada", "llegada_origen", "contacto_paciente_origen",
        "llegada_destino", "fecha_registro", "dt_registro",
        "fecha_hora_registro", "registro"
    ]
    for col in columnas_datetime:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    if "hora_programada" not in df.columns:
        df["hora_programada"] = pd.NaT
    if "fecha_programada" not in df.columns:
        df["fecha_programada"] = pd.NaT

    df["dt_programacion"] = df.apply(
        lambda r: combinar_fecha_hora(r.get("fecha_programada"), r.get("hora_programada")),
        axis=1
    )
    df["dt_registro_calc"] = df.apply(obtener_dt_registro, axis=1)
    return df


def obtener_codigo_regla(row):
    motivo = normalizar_texto(row.get("motivo_traslado"))
    sentido = normalizar_texto(row.get("sentido_traslado"))
    modalidad = normalizar_texto(row.get("modalidad"))

    if motivo == "CITA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return "Cita_Ida_Programadas"
    if motivo == "CITA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "Cita_Ida_No programado"
    if motivo == "CITA" and sentido == "RETORNO" and modalidad in ["PROGRAMADA", "NO PROGRAMADA"]:
        return "Cita_retorno_Programadas_No programadas"
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "PROGRAMADA":
        return "Referencias_Ida_Programadas"
    if motivo == "REFERENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "Referencias_Ida_No programado"
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "PROGRAMADA":
        return "Referencias_retorno_Programadas"
    if motivo == "REFERENCIA" and sentido == "RETORNO" and modalidad == "NO PROGRAMADA":
        return "Referencias_retorno_No programado"
    if motivo == "EMERGENCIA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "Emergencias_ida_No programado"
    if motivo == "ALTA" and sentido == "IDA" and modalidad == "NO PROGRAMADA":
        return "Altas_ida_No programado"
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
    }


def calcular_penalidades_fila(row):
    codigo = obtener_codigo_regla(row)
    tarifa = obtener_tarifa(row.get("tipo_unidad"))
    res = fila_resultado_base(codigo)

    dt_programacion = row.get("dt_programacion")
    dt_registro = row.get("dt_registro_calc")
    llegada_origen = row.get("llegada_origen")
    llegada_destino = row.get("llegada_destino")
    contacto_origen = row.get("contacto_paciente_origen")

    if codigo == "Cita_Ida_Programadas":
        if pd.notna(dt_programacion) and pd.notna(llegada_destino):
            hora_limite = dt_programacion - pd.Timedelta(minutes=15)
            atraso_real = minutos_diff(hora_limite, llegada_destino)
            res["diferencia_penalidad_destino"] = safe_round(atraso_real)
            if pd.notna(atraso_real) and atraso_real > 0:
                bloques = calcular_bloques(atraso_real, 30)
                res["min_penalidad_destino"] = safe_round(atraso_real)
                res["bloques_penalidad_destino"] = bloques
                res["penalidad_destino"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "CITA IDA PROGRAMADA"
                res["observacion_calculo"] = "Se compara llegada_destino con (hora_programada - 15 min)."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó dentro del tiempo permitido."
        if pd.notna(dt_programacion) and pd.notna(llegada_origen):
            anticipacion = minutos_diff(llegada_origen, dt_programacion)
            if pd.notna(anticipacion) and anticipacion > 60:
                res["min_penalidad_destino"] = 0.0
                res["bloques_penalidad_destino"] = 0
                res["penalidad_destino"] = 0.0
                res["observacion_calculo"] = "Penalidad eliminada: programación - llegada_origen > 60 min."
        return pd.Series(res)

    if codigo == "Cita_Ida_No programado":
        res["regla_aplicada"] = "CITA IDA NO PROGRAMADA"
        res["observacion_calculo"] = "No aplica penalidad según reglas de negocio."
        return pd.Series(res)

    if codigo == "Cita_retorno_Programadas_No programadas":
        if pd.notna(contacto_origen) and pd.notna(dt_programacion):
            atraso_real = minutos_diff(dt_programacion, contacto_origen)
            res["diferencia_penalidad_origen"] = safe_round(atraso_real)
            if pd.notna(atraso_real) and atraso_real > 0:
                bloques = calcular_bloques(atraso_real, 30)
                res["min_penalidad_origen"] = safe_round(atraso_real)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "CITA RETORNO PROGRAMADA/NO PROGRAMADA"
                res["observacion_calculo"] = "Se compara contacto_paciente_origen con hora_programada."
            else:
                res["observacion_calculo"] = "Sin penalidad: contacto dentro del horario."
        return pd.Series(res)

    if codigo == "Referencias_Ida_Programadas":
        if pd.isna(llegada_destino):
            res["perdida_cita_flag"] = 1
            res["perdida_cita_monto"] = MONTO_PERDIDA_CITA
            res["regla_aplicada"] = "REFERENCIA IDA PROGRAMADA"
            res["observacion_calculo"] = "Pérdida de cita: no existe llegada_destino."
            return pd.Series(res)
        if pd.notna(dt_programacion) and pd.notna(llegada_destino):
            atraso_real = minutos_diff(dt_programacion, llegada_destino)
            res["diferencia_penalidad_destino"] = safe_round(atraso_real)
            if pd.notna(llegada_origen):
                anticipacion = minutos_diff(llegada_origen, dt_programacion)
                if pd.notna(anticipacion) and anticipacion >= 90:
                    res["observacion_calculo"] = "Sin penalidad por exoneración: programación - llegada_origen >= 90 min."
                    return pd.Series(res)
            if pd.notna(atraso_real) and atraso_real > 0:
                bloques = calcular_bloques(atraso_real, 30)
                res["min_penalidad_destino"] = safe_round(atraso_real)
                res["bloques_penalidad_destino"] = bloques
                res["penalidad_destino"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "REFERENCIA IDA PROGRAMADA"
                res["observacion_calculo"] = "Se compara llegada_destino con hora_programada."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó dentro del horario."
        return pd.Series(res)

    if codigo == "Referencias_Ida_No programado":
        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)
            if pd.notna(demora_total) and demora_total > 30:
                exceso = demora_total - 30
                bloques = calcular_bloques(exceso, 30)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "REFERENCIA IDA NO PROGRAMADA"
                res["observacion_calculo"] = "Máximo 30 min desde registro hasta llegada_origen."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó a origen dentro de los 30 min desde el registro."
        return pd.Series(res)

    if codigo == "Referencias_retorno_Programadas":
        if pd.notna(contacto_origen) and pd.notna(dt_programacion):
            atraso_real = minutos_diff(dt_programacion, contacto_origen)
            res["diferencia_penalidad_origen"] = safe_round(atraso_real)
            if pd.notna(atraso_real) and atraso_real > 0:
                gracia = 30 if dt_programacion.year >= 2026 else 0
                atraso_penalizable = max(0, atraso_real - gracia)
                if atraso_penalizable > 0:
                    bloques = calcular_bloques(atraso_penalizable, 30)
                    res["min_penalidad_origen"] = safe_round(atraso_penalizable)
                    res["bloques_penalidad_origen"] = bloques
                    res["penalidad_origen"] = safe_round(bloques * tarifa)
                    res["regla_aplicada"] = "REFERENCIA RETORNO PROGRAMADA"
                    res["observacion_calculo"] = f"Se aplicó gracia de {gracia} min."
                else:
                    res["observacion_calculo"] = f"Sin penalidad: el atraso quedó cubierto por la gracia de {gracia} min."
            else:
                res["observacion_calculo"] = "Sin penalidad: contacto dentro del horario."
        return pd.Series(res)

    if codigo == "Emergencias_ida_No programado":
        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)
            barton = es_policlinico_barton(row)
            tolerancia = 15 if barton else 30
            bloque = 15 if barton else 30
            if pd.notna(demora_total) and demora_total > tolerancia:
                exceso = demora_total - tolerancia
                bloques = calcular_bloques(exceso, bloque)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "EMERGENCIA IDA NO PROGRAMADA"
                res["observacion_calculo"] = "Penalidad por llegada a origen fuera del tiempo máximo."
            else:
                res["observacion_calculo"] = "Sin penalidad."
        return pd.Series(res)

    if codigo == "Referencias_retorno_No programado":
        if pd.notna(llegada_origen) and pd.notna(dt_programacion):
            atraso_real = minutos_diff(dt_programacion, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(atraso_real)
            if pd.notna(atraso_real) and atraso_real > 0:
                gracia = 30 if dt_programacion.year >= 2026 else 0
                atraso_penalizable = max(0, atraso_real - gracia)
                if atraso_penalizable > 0:
                    bloques = calcular_bloques(atraso_penalizable, 30)
                    res["min_penalidad_origen"] = safe_round(atraso_penalizable)
                    res["bloques_penalidad_origen"] = bloques
                    res["penalidad_origen"] = safe_round(bloques * tarifa)
                    res["regla_aplicada"] = "REFERENCIA RETORNO NO PROGRAMADA"
                    res["observacion_calculo"] = f"Se aplicó gracia de {gracia} min."
                else:
                    res["observacion_calculo"] = f"Sin penalidad: el atraso quedó cubierto por la gracia de {gracia} min."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegada a origen dentro del horario."
        return pd.Series(res)

    if codigo == "Altas_ida_No programado":
        if pd.notna(dt_registro) and pd.notna(llegada_origen):
            demora_total = minutos_diff(dt_registro, llegada_origen)
            res["diferencia_penalidad_origen"] = safe_round(demora_total)
            if pd.notna(demora_total) and demora_total > 30:
                exceso = demora_total - 30
                bloques = calcular_bloques(exceso, 30)
                res["min_penalidad_origen"] = safe_round(exceso)
                res["bloques_penalidad_origen"] = bloques
                res["penalidad_origen"] = safe_round(bloques * tarifa)
                res["regla_aplicada"] = "ALTA IDA NO PROGRAMADA"
                res["observacion_calculo"] = "Máximo 30 min desde registro hasta llegada_origen."
            else:
                res["observacion_calculo"] = "Sin penalidad: llegó a origen dentro de los 30 min desde el registro."
        return pd.Series(res)

    res["observacion_calculo"] = "No existe regla configurada para esta combinación."
    return pd.Series(res)


def procesar_penalidades(df):
    df = normalizar_datos(df)
    resultado = df.apply(calcular_penalidades_fila, axis=1)
    df_salida = pd.concat([df, resultado], axis=1)
    df_salida["penalidad_total"] = (
        df_salida["penalidad_origen"].fillna(0)
        + df_salida["penalidad_destino"].fillna(0)
        + df_salida["perdida_cita_monto"].fillna(0)
    ).round(2)
    return df_salida


def exportar_excel_penalidades(df_resultado):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, sheet_name="penalidades", index=False)
    output.seek(0)
    return output
