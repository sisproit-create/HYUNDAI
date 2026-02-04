import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime

# Importa tu módulo (archivo original)
import hyundai_distribucion_equipos_v11 as hy

st.set_page_config(page_title="HYUNDAI | Distribución a Equipos", layout="wide")

# -------- Config ----------
HERE = Path(__file__).resolve().parent
DB_PATH = (HERE / hy.DB_PATH) if not Path(hy.DB_PATH).is_absolute() else Path(hy.DB_PATH)
TEMPLATE_XLSX = (HERE / hy.TEMPLATE_XLSX) if not Path(hy.TEMPLATE_XLSX).is_absolute() else Path(hy.TEMPLATE_XLSX)

# Asegura esquema al iniciar
hy.ensure_schema(str(DB_PATH))

# -------- Helpers ----------
def conn():
    return sqlite3.connect(str(DB_PATH))

def fetch_df(filters: dict) -> pd.DataFrame:
    with conn() as c:
        cols = hy.get_columns(c, "distribucion_hyundai_equipos")

        where = []
        params = []
        if filters.get("desde"):
            where.append("fecha >= ?")
            params.append(filters["desde"])
        if filters.get("hasta"):
            where.append("fecha <= ?")
            params.append(filters["hasta"])
        if filters.get("equipo"):
            where.append("equipo LIKE ?")
            params.append(f"%{filters['equipo'].strip().upper()}%")
        if filters.get("responsable"):
            where.append("responsable LIKE ?")
            params.append(f"%{filters['responsable'].strip()}%")
        if filters.get("tipo") and filters["tipo"] != "TODOS":
            where.append("tipo_registro = ?")
            params.append(filters["tipo"])

        sql = f"SELECT {', '.join(cols)} FROM distribucion_hyundai_equipos"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY fecha DESC, hora DESC, id DESC"

        df = pd.read_sql_query(sql, c, params=params)

    df = hy.normalize_df_columns(df)
    df = hy.aplicar_secuencia_contador(df)  # rellena secuencia si en BD hay NULL
    return df

def last_contador_final() -> float:
    with conn() as c:
        return hy.fetch_last_contador_final(c)

def last_horometro_final(equipo: str) -> float:
    with conn() as c:
        return hy.fetch_last_horometro_final(c, equipo)

def insert_record(data: dict) -> int:
    with conn() as c:
        return hy.insert_distribucion(c, data)

def update_record(rid: int, data: dict):
    with conn() as c:
        hy.update_distribucion(c, rid, data)

def delete_record(rid: int) -> bool:
    with conn() as c:
        return hy.delete_distribucion(c, rid)

def get_by_id(rid: int):
    with conn() as c:
        return hy.fetch_by_id(c, rid)

# -------- UI ----------
st.title("⛽ HYUNDAI | Distribución a Equipos (Dashboard)")
st.caption(f"DB: {DB_PATH.name}")

tab_reg, tab_list, tab_edit, tab_export, tab_tools = st.tabs(
    ["➕ Registrar", "📋 Listar", "✏️ Editar/Eliminar", "📦 Exportar", "🧰 Herramientas"]
)

# =======================
# TAB: Registrar
# =======================
with tab_reg:
    st.subheader("➕ Registrar distribución")

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        tipo_registro = st.selectbox(
            "Tipo de registro",
            ["HOROMETRO", "SIN_HOROMETRO"],
            key="reg_tipo_registro"
        )
    with col2:
        fecha = st.date_input(
            "Fecha",
            value=datetime.now().date(),
            key="reg_fecha"
        )
    with col3:
        hora = st.text_input(
            "Hora (HH:MM)",
            value=datetime.now().strftime("%H:%M"),
            key="reg_hora"
        )

    colA, colB, colC = st.columns([2, 1, 1])
    with colA:
        equipo = st.selectbox("Equipo", hy.EQUIPOS, key="reg_equipo")
        equipo_manual = st.text_input(
            "Si no está en la lista, escribe el equipo (opcional)",
            "",
            key="reg_equipo_manual"
        )
        equipo_final = (equipo_manual.strip().upper() if equipo_manual.strip() else equipo.strip().upper())
    with colB:
        responsable = st.selectbox("Responsable", hy.RESPONSABLES, key="reg_responsable")
    with colC:
        litros = st.number_input(
            "Litros despachados",
            min_value=0.0,
            step=1.0,
            format="%.2f",
            key="reg_litros"
        )

    gal = hy.litros_a_gal(float(litros))
    st.info(f"🔁 Conversión automática: **{litros:.2f} L = {gal:.2f} gal**")

    # Contadores (auto)
    st.markdown("### 📟 Contadores (auto)")
    ci_default = last_contador_final()
    colx, coly = st.columns([1, 1])
    with colx:
        contador_inicial = st.number_input(
            "Contador inicial (L)",
            value=float(ci_default),
            step=1.0,
            format="%.2f",
            key="reg_contador_inicial"
        )
    contador_final_calc = round(float(contador_inicial) + float(litros), 2)
    with coly:
        contador_final = st.number_input(
            "Contador final (L)",
            value=float(contador_final_calc),
            step=1.0,
            format="%.2f",
            key="reg_contador_final"
        )

    # Horómetro si aplica
    horometro_inicial = None
    horometro_final = None
    horas_trabajadas = None
    consumo_por_gl_h = None

    if tipo_registro == "HOROMETRO":
        st.markdown("### ⏱️ Horómetro")
        hi_default = last_horometro_final(equipo_final)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            horometro_inicial = st.number_input(
                "Horómetro inicial",
                value=float(hi_default),
                step=0.1,
                format="%.2f",
                key="reg_horometro_inicial"
            )
        with c2:
            horometro_final = st.number_input(
                "Horómetro final",
                value=float(hi_default),
                step=0.1,
                format="%.2f",
                key="reg_horometro_final"
            )
        horas_trabajadas = round(float(horometro_final) - float(horometro_inicial), 2)
        if horas_trabajadas <= 0:
            st.error("⚠️ Horas trabajadas debe ser > 0 (horómetro final > inicial).")
        else:
            consumo_por_gl_h = round(float(gal) / float(horas_trabajadas), 2)
            with c3:
                st.metric("Horas trabajadas", f"{horas_trabajadas:.2f}")
            st.metric("Consumo (gal/h)", f"{consumo_por_gl_h:.2f}")

    # Precio diesel
    st.markdown("### 💲 Precio y costo")
    precio_default = hy.obtener_precio_diesel_actual()
    colp, colc = st.columns([1, 1])
    with colp:
        precio_diesel = st.number_input(
            "Precio diésel (USD/gal)",
            value=float(precio_default),
            step=0.0001,
            format="%.4f",
            key="reg_precio_diesel"
        )
    costo = round(float(gal) * float(precio_diesel), 2)
    with colc:
        st.metric("Costo estimado (USD)", f"{costo:.2f}")

    st.divider()
    if st.button("✅ Guardar registro", type="primary", key="reg_guardar"):
        if tipo_registro == "HOROMETRO" and (horas_trabajadas is None or horas_trabajadas <= 0):
            st.stop()

        data = dict(
            fecha=fecha.strftime("%Y-%m-%d"),
            hora=hora.strip(),
            equipo=equipo_final,
            responsable=responsable,
            litros_despachados=float(litros),
            volumen_despachado=float(gal),
            contador_inicial=float(contador_inicial),
            contador_final=float(contador_final),
            horometro_inicial=(float(horometro_inicial) if horometro_inicial is not None else None),
            horometro_final=(float(horometro_final) if horometro_final is not None else None),
            horas_trabajadas=(float(horas_trabajadas) if horas_trabajadas is not None else None),
            consumo_por_gl_h=(float(consumo_por_gl_h) if consumo_por_gl_h is not None else None),
            precio_diesel=float(precio_diesel),
            costo_diesel_usd=float(costo),
            volumen_restante_hyundai=None,
            tipo_registro=tipo_registro,
        )
        new_id = insert_record(data)
        st.success(f"Guardado ✅ ID = {new_id}")

# =======================
# TAB: Listar
# =======================
with tab_list:
    st.subheader("📋 Listar registros")

    f1, f2, f3, f4, f5 = st.columns([1, 1, 1.5, 1.5, 1])
    with f1:
        desde = st.text_input("Desde (YYYY-MM-DD)", value="", key="list_desde")
    with f2:
        hasta = st.text_input("Hasta (YYYY-MM-DD)", value="", key="list_hasta")
    with f3:
        equipo_f = st.text_input("Equipo contiene", value="", key="list_equipo")
    with f4:
        responsable_f = st.text_input("Responsable contiene", value="", key="list_responsable")
    with f5:
        tipo_f = st.selectbox("Tipo", ["TODOS", "HOROMETRO", "SIN_HOROMETRO"], key="list_tipo")

    df = fetch_df(dict(desde=desde, hasta=hasta, equipo=equipo_f, responsable=responsable_f, tipo=tipo_f))

    st.caption(f"Registros: {len(df)}")
    show_cols = [c for c in [
        "id","fecha","hora","equipo","litros_despachados","volumen_despachado","responsable",
        "contador_inicial_show","contador_final_show","Secuencia_Contador","Delta_Litros","tipo_registro",
        "horometro_inicial","horometro_final","horas_trabajadas","consumo_por_gl_h","precio_diesel","costo_diesel_usd"
    ] if c in df.columns]

    st.dataframe(df[show_cols], use_container_width=True, height=520)

    csv = df[show_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Descargar CSV",
        data=csv,
        file_name="hyundai_distribucion.csv",
        mime="text/csv",
        key="list_download_csv"
    )

# =======================
# TAB: Editar / Eliminar
# =======================
with tab_edit:
    st.subheader("✏️ Editar / 🗑️ Eliminar")

    rid = st.number_input("ID", min_value=1, step=1, key="edit_rid")
    rec = get_by_id(int(rid)) if rid else None

    if not rec:
        st.info("Escribe un ID y se cargará el registro.")
    else:
        st.success("Registro cargado.")

        # Keys atadas al rid para que al cambiar el ID se refresquen los valores
        kpref = f"edit_{int(rid)}"

        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            e_fecha = st.text_input("Fecha", value=str(rec.get("fecha", "")), key=f"{kpref}_fecha")
        with c2:
            e_hora = st.text_input("Hora", value=str(rec.get("hora", "")), key=f"{kpref}_hora")
        with c3:
            e_tipo = st.selectbox(
                "Tipo registro",
                ["HOROMETRO", "SIN_HOROMETRO"],
                index=0 if rec.get("tipo_registro") == "HOROMETRO" else 1,
                key=f"{kpref}_tipo"
            )

        c4, c5 = st.columns([2, 1])
        with c4:
            e_equipo = st.text_input("Equipo", value=str(rec.get("equipo", "")), key=f"{kpref}_equipo")
        with c5:
            e_resp = st.text_input("Responsable", value=str(rec.get("responsable", "")), key=f"{kpref}_responsable")

        c6, c7, c8 = st.columns([1, 1, 1])
        with c6:
            e_litros = st.number_input(
                "Litros",
                value=float(rec.get("litros_despachados") or 0.0),
                step=1.0,
                format="%.2f",
                key=f"{kpref}_litros"
            )
        e_gal = hy.litros_a_gal(float(e_litros))
        with c7:
            st.metric("Galones", f"{e_gal:.2f}")
        with c8:
            e_precio = st.number_input(
                "Precio diesel",
                value=float(rec.get("precio_diesel") or hy.obtener_precio_diesel_actual()),
                step=0.0001,
                format="%.4f",
                key=f"{kpref}_precio"
            )

        c9, c10 = st.columns([1, 1])
        with c9:
            e_ci = st.number_input(
                "Contador inicial",
                value=float(rec.get("contador_inicial") or 0.0),
                step=1.0,
                format="%.2f",
                key=f"{kpref}_cont_ini"
            )
        cf_sug = round(float(e_ci) + float(e_litros), 2)
        with c10:
            e_cf = st.number_input(
                "Contador final",
                value=float(rec.get("contador_final") or cf_sug),
                step=1.0,
                format="%.2f",
                key=f"{kpref}_cont_fin"
            )

        # Horómetro editable
        st.markdown("### ⏱️ Horómetro (si aplica)")
        h1, h2, h3 = st.columns([1, 1, 1])
        with h1:
            e_hi = st.number_input(
                "Horómetro inicial",
                value=float(rec.get("horometro_inicial") or 0.0),
                step=0.1,
                format="%.2f",
                key=f"{kpref}_horo_ini"
            )
        with h2:
            e_hf = st.number_input(
                "Horómetro final",
                value=float(rec.get("horometro_final") or 0.0),
                step=0.1,
                format="%.2f",
                key=f"{kpref}_horo_fin"
            )

        horas = round(float(e_hf) - float(e_hi), 2)
        consumo = None
        if horas > 0:
            consumo = round(float(e_gal) / horas, 2)

        with h3:
            st.metric("Horas trabajadas", f"{horas:.2f}")

        costo = round(float(e_gal) * float(e_precio), 2)
        st.metric("Costo diesel (USD)", f"{costo:.2f}")

        col_save, col_del = st.columns([1, 1])
        with col_save:
            if st.button("💾 Guardar cambios", type="primary", key=f"{kpref}_guardar"):
                update_record(int(rid), dict(
                    fecha=e_fecha.strip(),
                    hora=e_hora.strip(),
                    equipo=e_equipo.strip().upper(),
                    responsable=e_resp.strip(),
                    litros_despachados=float(e_litros),
                    volumen_despachado=float(e_gal),
                    contador_inicial=float(e_ci),
                    contador_final=float(e_cf),
                    horometro_inicial=float(e_hi) if e_tipo == "HOROMETRO" else None,
                    horometro_final=float(e_hf) if e_tipo == "HOROMETRO" else None,
                    horas_trabajadas=float(horas) if (e_tipo == "HOROMETRO" and horas > 0) else None,
                    consumo_por_gl_h=float(consumo) if (e_tipo == "HOROMETRO" and consumo is not None) else None,
                    precio_diesel=float(e_precio),
                    costo_diesel_usd=float(costo),
                    tipo_registro=e_tipo,
                ))
                st.success("Actualizado ✅")

        with col_del:
            if st.button("🗑️ Eliminar registro", type="secondary", key=f"{kpref}_eliminar"):
                ok = delete_record(int(rid))
                st.warning("Eliminado ✅" if ok else "No se pudo eliminar")

# =======================
# TAB: Exportar
# =======================
with tab_export:
    st.subheader("📦 Exportar mes a Excel (plantilla)")

    st.write("Plantilla esperada:", str(TEMPLATE_XLSX))
    mes = st.text_input("Mes (YYYY-MM)", value=datetime.now().strftime("%Y-%m"), key="exp_mes")

    if st.button("📤 Exportar", key="exp_exportar"):
        if not mes or len(mes) != 7 or mes[4] != "-":
            st.error("Formato inválido. Usa YYYY-MM")
        elif not TEMPLATE_XLSX.exists():
            st.error("No encuentro la plantilla .xlsx. Colócala en la misma carpeta del app.")
        else:
            with conn() as c:
                cols = hy.get_columns(c, "distribucion_hyundai_equipos")
                sql = f"SELECT {', '.join(cols)} FROM distribucion_hyundai_equipos WHERE fecha LIKE ? ORDER BY fecha ASC, hora ASC, id ASC"
                df = pd.read_sql_query(sql, c, params=[f"{mes}-%"])

            df = hy.normalize_df_columns(df)
            df = hy.aplicar_secuencia_contador(df)

            if df.empty:
                st.warning("No hay registros para ese mes.")
            else:
                out_name = HERE / f"Distribucion_Equipos_{mes}.xlsx"
                df.to_excel(out_name, index=False)
                with open(out_name, "rb") as f:
                    st.download_button(
                        "⬇️ Descargar Excel",
                        data=f,
                        file_name=out_name.name,
                        key="exp_download_excel"
                    )

# =======================
# TAB: Tools
# =======================
with tab_tools:
    st.subheader("🧰 Herramientas")

    if st.button("🧱 Backfill contadores (llenar NULL)", key="tools_backfill"):
        hy.backfill_contadores(str(DB_PATH))
        st.success("Backfill ejecutado ✅ (revisa listado).")
