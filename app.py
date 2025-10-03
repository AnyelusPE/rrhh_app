import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="RRHH - Control de Tardanzas", layout="wide")

st.title("📊 RRHH - Versión Completa")

# Subir archivos
file_marc = st.file_uploader("Sube archivo de Marcaciones (Excel)", type=["xlsx", "xls"])
file_hor = st.file_uploader("Sube archivo de Horarios (Excel)", type=["xlsx", "xls"])

if st.button("Procesar Datos"):
    if not file_marc or not file_hor:
        st.error("⚠️ Debes subir ambos archivos (marcaciones y horarios).")
    else:
        try:
            # --------------------
            # LEER MARCACIONES
            # --------------------
            df_marc = pd.read_excel(file_marc, dtype=str, engine="openpyxl")
            df_marc.columns = df_marc.columns.astype(str).str.strip().str.upper()

            # Renombrar columnas para estandarizar
            renombres = {"NO.": "DNI"}
            for col in df_marc.columns:
                if "DEPART" in col:  # Acepta DEPARTAMENTO, DEPARTMENT
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
            df_hor.columns = df_hor.columns.astype(str).str.strip().str.upper()

            columnas_hor_base = ["DNI", "NOMBRE Y APELLIDO", "ID"]
            for col in columnas_hor_base:
                if col not in df_hor.columns:
                    st.error(f"Falta la columna obligatoria en horarios: {col}. Encontradas: {list(df_hor.columns)}")
                    st.stop()

            # Detectar columnas de fechas
            columnas_fechas = []
            for col in df_hor.columns:
                try:
                    fecha = pd.to_datetime(str(col), dayfirst=True, errors="raise")
                    columnas_fechas.append(col)
                except:
                    continue

            if not columnas_fechas:
                st.error("No se encontraron columnas de fechas válidas en horarios")
                st.stop()

            # Transformar horarios a formato largo
            horarios_largos = []
            for fecha_col in columnas_fechas:
                fecha = pd.to_datetime(str(fecha_col), dayfirst=True).date()
                temp_df = df_hor[["DNI", "NOMBRE Y APELLIDO", "ID"]].copy()
                temp_df["FECHA"] = fecha
                temp_df["HORARIO_ESPERADO"] = df_hor[fecha_col].astype(str).str.strip()
                horarios_largos.append(temp_df)

            df_hor_largo = pd.concat(horarios_largos, ignore_index=True)

            # --------------------
            # CRUCE MARCACIONES VS HORARIOS
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

                resultado.append(
                    {
                        "DNI": dni,
                        "NOMBRE": row["NOMBRE Y APELLIDO"],
                        "FECHA": fecha,
                        "HORARIO": horario,
                        "TARDANZA (min)": tardanza,
                    }
                )

            df_resultado = pd.DataFrame(resultado)

            # --------------------
            # PIVOT HORIZONTAL
            # --------------------
            df_pivot = df_resultado.pivot_table(
                index=["DNI", "NOMBRE"],
                columns="FECHA",
                values="TARDANZA (min)",
                aggfunc="first"
            ).reset_index()

            # --------------------
            # MOSTRAR Y DESCARGAR
            # --------------------
            st.success("✅ Procesamiento completado")
            st.write("### Resultado de Tardanzas (Formato Horizontal)")
            st.dataframe(df_pivot)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_pivot.to_excel(writer, index=False, sheet_name="Tardanzas")
            st.download_button(
                label="📥 Descargar Resultado en Excel",
                data=buffer,
                file_name="tardanzas_resumen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error durante el procesamiento: {str(e)}")