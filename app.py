import io
import re
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st


st.set_page_config(page_title="RRHH - Control de Tardanzas", page_icon="🕒", layout="wide")

st.title("🕒 Sistema de Marcaciones")


@st.cache_data(show_spinner=False)
def read_excel_bytes(uploaded_file) -> bytes:
    return uploaded_file.getvalue() if uploaded_file else b""


def normaliza_columnas(cols):
    out = []
    for c in cols:
        if isinstance(c, pd.Timestamp):
            out.append(c.date().strftime("%Y-%m-%d"))
        else:
            out.append(str(c).strip().upper())
    return pd.Index(out)


def leer_marcaciones(file) -> pd.DataFrame:
    data = read_excel_bytes(file)
    df = pd.read_excel(io.BytesIO(data), dtype=str)
    df.columns = df.columns.astype(str).str.strip().str.upper()

    renombres = {"NO.": "DNI"}
    for col in df.columns:
        if "DEPART" in col:
            renombres[col] = "DEPARTAMENTO"
    df.rename(columns=renombres, inplace=True)

    requeridas = ["DEPARTAMENTO", "NOMBRE", "DNI", "FECHA/HORA", "ESTADO"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en marcaciones: {faltantes}. Encontradas: {list(df.columns)}")

    df["FECHA/HORA"] = pd.to_datetime(df["FECHA/HORA"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["FECHA/HORA"]).copy()
    df["FECHA"] = df["FECHA/HORA"].dt.date
    return df


def leer_horarios(file) -> pd.DataFrame:
    data = read_excel_bytes(file)
    df = pd.read_excel(io.BytesIO(data), dtype=str)
    df.columns = normaliza_columnas(df.columns)

    requeridas = ["DNI", "NOMBRE Y APELLIDO", "ID"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en horarios: {faltantes}. Encontradas: {list(df.columns)}")

    columnas_fechas = []
    for col in df.columns:
        try:
            pd.to_datetime(col, dayfirst=True, errors="raise")
            columnas_fechas.append(col)
        except Exception:
            continue

    if not columnas_fechas:
        raise ValueError("No se detectaron columnas de fecha en el archivo de horarios.")

    df_long = df.melt(
        id_vars=["DNI", "NOMBRE Y APELLIDO", "ID"],
        value_vars=columnas_fechas,
        var_name="FECHA",
        value_name="HORARIO_ESPERADO",
    )
    df_long["FECHA"] = pd.to_datetime(df_long["FECHA"], dayfirst=True, errors="coerce").dt.date
    df_long["HORARIO_ESPERADO"] = df_long["HORARIO_ESPERADO"].astype(str).str.strip()
    return df_long


def extrae_hora_inicio(horario: str) -> str | None:
    if not horario or str(horario).strip() == "" or "DESCANSO" in str(horario).upper():
        return None
    m = re.search(r"(\d{1,2}:\d{2})", str(horario))
    return m.group(1) if m else None


def horas_a_minutos(hhmm: str | None) -> float | None:
    if not hhmm:
        return None
    try:
        h, m = hhmm.split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return None


def calcula_tardanzas(df_marc: pd.DataFrame, df_hor: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # Primera marcación por DNI/FECHA
    df_first = (
        df_marc.sort_values("FECHA/HORA")
        .groupby(["DNI", "FECHA"], as_index=False)["FECHA/HORA"].first()
        .rename(columns={"FECHA/HORA": "PRIMERA_MARCACION"})
    )
    df_first["PRIMERA_MIN"] = df_first["PRIMERA_MARCACION"].dt.hour * 60 + df_first["PRIMERA_MARCACION"].dt.minute

    df_hor = df_hor.copy()
    df_hor["HORA_INICIO_STR"] = df_hor["HORARIO_ESPERADO"].apply(extrae_hora_inicio)
    df_hor["HORA_INICIO_MIN"] = df_hor["HORA_INICIO_STR"].apply(horas_a_minutos)
    df_hor["DESCANSO"] = df_hor["HORA_INICIO_STR"].isna()

    base = df_hor.merge(df_first, on=["DNI", "FECHA"], how="left")

    def tardanza_row(row):
        if row["DESCANSO"]:
            return 0
        if pd.isna(row.get("HORA_INICIO_MIN")):
            return "-"
        if pd.isna(row.get("PRIMERA_MIN")):
            return "-"
        delta = int(row["PRIMERA_MIN"] - row["HORA_INICIO_MIN"])
        return max(0, delta)

    base["TARDANZA (min)"] = base.apply(tardanza_row, axis=1)
    df_resultado = base[[
        "DNI",
        "NOMBRE Y APELLIDO",
        "FECHA",
        "HORARIO_ESPERADO",
        "TARDANZA (min)",
    ]].rename(columns={"NOMBRE Y APELLIDO": "NOMBRE", "HORARIO_ESPERADO": "HORARIO"})

    # Pivot + total
    df_pivot = df_resultado.pivot_table(
        index=["DNI", "NOMBRE"],
        columns="FECHA",
        values="TARDANZA (min)",
        aggfunc="first",
    ).reset_index()

    cols_val = [c for c in df_pivot.columns if c not in ("DNI", "NOMBRE")]
    def safe_sum(row):
        s = 0
        for x in row:
            if isinstance(x, (int, float)):
                s += x
        return s
    df_pivot["TOTAL_TARDANZA"] = df_pivot[cols_val].apply(safe_sum, axis=1)

    return df_resultado, df_pivot


def calcula_horas(df_marc: pd.DataFrame) -> pd.DataFrame:
    def format_td(td: timedelta) -> str:
        total_seconds = int(td.total_seconds())
        if total_seconds < 0:
            total_seconds = 0
        h = total_seconds // 3600
        m = (total_seconds % 3600) // 60
        s = total_seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    filas = []
    for (dni, fecha), grupo in df_marc.groupby(["DNI", "FECHA"], sort=True):
        g = grupo.sort_values("FECHA/HORA").reset_index(drop=True)
        entrada = g.loc[0, "FECHA/HORA"]
        salida = g.loc[len(g) - 1, "FECHA/HORA"]

        if len(g) >= 4:
            inicio_refri = g.loc[1, "FECHA/HORA"]
            fin_refri = g.loc[2, "FECHA/HORA"]
            dur_refri = fin_refri - inicio_refri
        else:
            inicio_refri = None
            fin_refri = None
            dur_refri = timedelta(0)

        horas_trab = salida - entrada - dur_refri

        filas.append({
            "DNI": dni,
            "NOMBRE": g.loc[0, "NOMBRE"],
            "FECHA": fecha,
            "ENTRADA": entrada.strftime("%H:%M:%S"),
            "SALIDA": salida.strftime("%H:%M:%S"),
            "HORAS_TRABAJADAS": format_td(horas_trab),
            "INICIO_REFRIGERIO": inicio_refri.strftime("%H:%M:%S") if inicio_refri else "-",
            "FIN_REFRIGERIO": fin_refri.strftime("%H:%M:%S") if fin_refri else "-",
            "DURACION_REFRIGERIO": format_td(dur_refri),
        })

    return pd.DataFrame(filas)


# Carga de archivos
file_marc = st.file_uploader("Sube archivo de Marcaciones (Excel)", type=["xlsx", "xls"])
file_hor = st.file_uploader("Sube archivo de Horarios (Excel)", type=["xlsx", "xls"])

if st.button("Procesar Datos"):
    if not file_marc or not file_hor:
        st.error("Debes subir ambos archivos (marcaciones y horarios).")
    else:
        with st.spinner("Procesando cruces de marcaciones y horarios..."):
            try:
                df_marc = leer_marcaciones(file_marc)
                df_hor = leer_horarios(file_hor)

                df_resultado, df_pivot = calcula_tardanzas(df_marc, df_hor)
                df_horas = calcula_horas(df_marc)

                st.success("Procesamiento completado correctamente.")

                st.write("### Tardanzas por día")
                st.dataframe(df_pivot, use_container_width=True)

                st.write("### Horas Trabajadas y Refrigerio")
                st.dataframe(df_horas, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_pivot.to_excel(writer, index=False, sheet_name="Tardanzas")
                    df_horas.to_excel(writer, index=False, sheet_name="Horas_Trabajadas")
                buffer.seek(0)

                st.download_button(
                    label="Descargar resultado en Excel",
                    data=buffer,
                    file_name="reporte_rrhh_completo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Error durante el procesamiento: {e}")
