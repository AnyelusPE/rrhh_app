import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="RRHH - Control de Tardanzas", layout="wide")

st.title("📊 SISTEMA DE MARCACIONES")

# Subir archivos
file_marc = st.file_uploader("Sube archivo de Marcaciones (Excel)", type=["xlsx", "xls"])
file_hor = st.file_uploader("Sube archivo de Horarios (Excel)", type=["xlsx", "xls"])

if st.button("Procesar Datos"):
    if not file_marc or not file_hor:
        st.error("⚠️ Debes subir ambos archivos (marcaciones y horarios).")
    else:
        with st.spinner("⏳ Procesando cruces de marcaciones y horarios..."):
            try:
                # --------------------
                # LEER MARCACIONES
                # --------------------
                df_marc = pd.read_excel(file_marc, dtype=str, engine="openpyxl")
                df_marc.columns = df_marc.columns.astype(str).str.strip().str.upper()

                renombres = {"NO.": "DNI"}
                for col in df_marc.columns:
                    if "DEPART" in col:
                        renombres[col] = "DEPARTAMENTO"
                df_marc.rename(columns=renombres, inplace=True)

                columnas_marc_base = ["DEPARTAMENTO", "NOMBRE", "DNI", "FECHA/HORA", "ESTADO"]
                for col in columnas_marc_base:
                    if col not in df_marc.columns:
                        st.error(f"Falta la columna obligatoria en marcaciones: {col}. Encontradas: {list(df_marc.columns)}")
                        st.stop()

                df_marc["FECHA/HORA"] = pd.to_datetime(df_marc["FECHA/HORA"], errors="coerce", dayfirst=True)
                df_marc = df_marc.dropna(subset=["FECHA/HORA"])
                df_marc["FECHA"] = df_marc["FECHA/HORA"].dt.date
                df_marc["HORA"] = df_marc["FECHA/HORA"].dt.time

                # --------------------
                # LEER HORARIOS
                # --------------------
                df_hor = pd.read_excel(file_hor, dtype=str, engine="openpyxl")
                df_hor.columns = [c if not isinstance(c, pd.Timestamp) else c.date().strftime("%Y-%m-%d") for c in df_hor.columns]
                df_hor.columns = pd.Index([str(c).strip().upper() for c in df_hor.columns])

                columnas_hor_base = ["DNI", "NOMBRE Y APELLIDO", "ID"]
                for col in columnas_hor_base:
                    if col not in df_hor.columns:
                        st.error(f"Falta la columna obligatoria en horarios: {col}. Encontradas: {list(df_hor.columns)}")
                        st.stop()

                # Detectar columnas de fechas
                columnas_fechas = []
                for col in df_hor.columns:
                    try:
                        fecha = pd.to_datetime(col, dayfirst=True, errors="raise")
                        columnas_fechas.append(col)
                    except:
                        continue

                # --------------------
                # FORMATO LARGO
                # --------------------
                horarios_largos = []
                for fecha_col in columnas_fechas:
                    fecha = pd.to_datetime(str(fecha_col), dayfirst=True).date()
                    temp_df = df_hor[["DNI", "NOMBRE Y APELLIDO", "ID"]].copy()
                    temp_df["FECHA"] = fecha
                    temp_df["HORARIO_ESPERADO"] = df_hor[fecha_col].astype(str).str.strip()
                    horarios_largos.append(temp_df)

                df_hor_largo = pd.concat(horarios_largos, ignore_index=True)

                # --------------------
                # CRUCE DE TARDANZAS
                # --------------------
                resultado = []
                for _, row in df_hor_largo.iterrows():
                    dni = str(row["DNI"]).strip()
                    fecha = row["FECHA"]
                    horario = str(row["HORARIO_ESPERADO"]).upper()

                    if "DESCANSO" in horario or horario == "" or pd.isna(horario):
                        tardanza = 0
                    else:
                        try:
                            hora_inicio = horario.split("-")[0].strip()
                            hora_inicio_dt = datetime.strptime(hora_inicio, "%H:%M").time()

                            registros = df_marc[(df_marc["DNI"] == dni) & (df_marc["FECHA"] == fecha)]
                            if not registros.empty:
                                primera_marcacion = registros["FECHA/HORA"].min().time()
                                delta = (
                                    datetime.combine(datetime.today(), primera_marcacion)
                                    - datetime.combine(datetime.today(), hora_inicio_dt)
                                ).total_seconds() / 60
                                tardanza = max(0, int(delta))
                            else:
                                tardanza = "-"
                        except:
                            tardanza = "-"

                    resultado.append({
                        "DNI": dni,
                        "NOMBRE": row["NOMBRE Y APELLIDO"],
                        "FECHA": fecha,
                        "HORARIO": horario,
                        "TARDANZA (min)": tardanza,
                    })

                df_resultado = pd.DataFrame(resultado)

                # --------------------
                # PIVOT + TOTAL
                # --------------------
                df_pivot = df_resultado.pivot_table(
                    index=["DNI", "NOMBRE"],
                    columns="FECHA",
                    values="TARDANZA (min)",
                    aggfunc="first"
                ).reset_index()

                df_pivot["TOTAL_TARDANZA"] = df_pivot.drop(columns=["DNI", "NOMBRE"]).apply(
                    lambda row: sum([x if isinstance(x, (int, float)) else 0 for x in row]), axis=1
                )

                # --------------------
                # HORAS TRABAJADAS Y REFRIGERIO
                # --------------------
                horas = []
                for (dni, fecha), grupo in df_marc.groupby(["DNI", "FECHA"]):
                    grupo_ordenado = grupo.sort_values("FECHA/HORA")

                    entrada = grupo_ordenado["FECHA/HORA"].iloc[0]
                    salida = grupo_ordenado["FECHA/HORA"].iloc[-1]

                    # Horas intermedias (refrigerio)
                    if len(grupo_ordenado) >= 4:
                        inicio_refri = grupo_ordenado["FECHA/HORA"].iloc[1]
                        fin_refri = grupo_ordenado["FECHA/HORA"].iloc[2]
                        dur_refri = (fin_refri - inicio_refri)
                    else:
                        inicio_refri, fin_refri, dur_refri = None, None, timedelta(0)

                    horas_trab = (salida - entrada - dur_refri)

                    horas.append({
                        "DNI": dni,
                        "NOMBRE": grupo_ordenado["NOMBRE"].iloc[0],
                        "FECHA": fecha,
                        "ENTRADA": entrada.strftime("%H:%M:%S"),
                        "SALIDA": salida.strftime("%H:%M:%S"),
                        "HORAS_TRABAJADAS": str(horas_trab).split(".")[0],
                        "INICIO_REFRIGERIO": inicio_refri.strftime("%H:%M:%S") if inicio_refri else "-",
                        "FIN_REFRIGERIO": fin_refri.strftime("%H:%M:%S") if fin_refri else "-",
                        # 🔧 Corregido: quitar "0 days" del resultado
                        "DURACIÓN_REFRIGERIO": str(dur_refri).replace("0 days ", "").split(".")[0] if dur_refri != timedelta(0) else "00:00:00",
                    })

                df_horas = pd.DataFrame(horas)

                # --------------------
                # DESCARGA
                # --------------------
                st.success("✅ Procesamiento completado correctamente.")
                st.write("### 📅 Resultado de Tardanzas")
                st.dataframe(df_pivot, use_container_width=True)

                st.write("### ⏱️ Resultado de Horas Trabajadas y Refrigerio")
                st.dataframe(df_horas, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_pivot.to_excel(writer, index=False, sheet_name="Tardanzas")
                    df_horas.to_excel(writer, index=False, sheet_name="Horas_Trabajadas")

                st.download_button(
                    label="📥 Descargar Resultado en Excel",
                    data=buffer,
                    file_name="reporte_rrhh_completo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {str(e)}")
