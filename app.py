import json
import os
import time

import pandas as pd
import requests
import streamlit as st
from sqlalchemy import text
from sqlalchemy.orm import Session

from base import engine

# --- CONFIGURACIÓN DE RUTAS Y API ---
API_URL = "http://127.0.0.1:8000/api/procesar-factura"
ARCHIVO_EXCEL = "reporte_extracciones.xlsx"

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Extractor IA Automotriz",
    page_icon=":material/directions_car:",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .main-header { font-size: 2.2rem; font-weight: 600; color: #9c0094; margin-bottom: 0px; padding-bottom: 0px;}
    .sub-header { font-size: 1.1rem; color: #64748B; margin-bottom: 2rem; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Sistema de Gestión y Extracción de Facturas</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Motor de IA Centralizado mediante API REST</div>', unsafe_allow_html=True)

tab_facturas, tab_extracciones, tab_vehiculos = st.tabs([
    ":material/upload_file: Procesamiento por Lotes",
    ":material/table_chart: Reporte de Extracciones",
    ":material/database: Base de Datos de Vehículos"
])


# ==========================================
# PESTAÑA 1: SUBIR FACTURAS Y VER DESGLOSE (TIEMPO REAL)
# ==========================================
with tab_facturas:
    with st.container(border=True):
        st.subheader(":material/memory: Motor de Procesamiento IA (Remoto)")
        st.info("Sube documentos de múltiples páginas. Verás los resultados aparecer uno por uno en tiempo real.", icon=":material/info:")

        archivos_subidos = st.file_uploader(
            "Selecciona los documentos (PDF, PNG, JPG)",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            label_visibility="collapsed"
        )

        if archivos_subidos:
            st.write("")
            if st.button("Enviar al Servidor IA", type="primary", icon=":material/cloud_upload:"):
                total_archivos = len(archivos_subidos)
                procesados = 0
                barra_progreso = st.progress(0, text="Preparando envío...")

                for archivo in archivos_subidos:
                    st.markdown(f"### 📄 Desglose de **{archivo.name}**")
                    st.toast(f"Procesando {archivo.name}...", icon=":material/sync:")

                    try:
                        files = {"file": (archivo.name, archivo.getvalue(), archivo.type or "application/octet-stream")}
                        with requests.post(API_URL, files=files, stream=True, timeout=None) as respuesta:
                            if respuesta.status_code == 200:
                                idx = 1
                                recibio_algo = False
                                for linea in respuesta.iter_lines(decode_unicode=True):
                                    if not linea:
                                        continue
                                    try:
                                        doc = json.loads(linea)
                                    except json.JSONDecodeError:
                                        st.warning(f"Respuesta no JSON del servidor: {linea[:240]}")
                                        continue

                                    recibio_algo = True
                                    estado = str(doc.get("estado", "DESCONOCIDO"))
                                    if "ÉXITO" in estado:
                                        simbolo, color_estado = "🟢", "green"
                                    else:
                                        simbolo, color_estado = "🟠", "darkorange"

                                    titulo_expander = (
                                        f"{simbolo} Factura Segmentada {idx} | "
                                        f"{doc.get('marca_base', 'N/A')} {doc.get('modelo_base', '')}"
                                    )

                                    with st.expander(titulo_expander, expanded=True):
                                        c1, c2, c3, c4 = st.columns(4)
                                        c1.metric("Marca Detectada", doc.get("marca") or "-")
                                        c2.metric("Modelo Detectado", doc.get("modelo") or "-")
                                        c3.metric("Beneficiario", str(doc.get("beneficiario") or "-")[:15])
                                        costo = doc.get("costo_total", 0) or 0
                                        c4.metric("Costo Total", f"${costo}")

                                        paginas = doc.get("paginas_pdf") or []
                                        if paginas:
                                            st.caption(f"Páginas PDF: {paginas}")

                                        st.markdown(
                                            f"**Estado:** <span style='color:{color_estado}; font-weight:bold;'>{estado}</span>",
                                            unsafe_allow_html=True
                                        )
                                        st.json(doc)

                                    idx += 1

                                if not recibio_algo:
                                    st.warning(f"El servidor no devolvió facturas para '{archivo.name}'.")
                            else:
                                st.error(
                                    f"Error en servidor al procesar '{archivo.name}': {respuesta.text}",
                                    icon=":material/warning:"
                                )

                    except requests.exceptions.ConnectionError:
                        st.error(
                            f"No se pudo conectar a la API ({API_URL}). ¿Encendiste el servidor `api_backend.py`?",
                            icon=":material/cloud_off:"
                        )
                        break
                    except Exception as e:
                        st.error(f"Error inesperado con '{archivo.name}': {e}", icon=":material/error:")

                    procesados += 1
                    barra_progreso.progress(
                        procesados / total_archivos,
                        text=f"Procesando: {procesados} de {total_archivos} archivos completados"
                    )
                    st.divider()

                if procesados == total_archivos:
                    st.success("Procesamiento por lotes finalizado exitosamente.", icon=":material/task_alt:")


# ==========================================
# PESTAÑA 2: VER EXTRACCIONES
# ==========================================
with tab_extracciones:
    col_header1, col_header2 = st.columns([8, 2], vertical_alignment="bottom")
    with col_header1:
        st.subheader(":material/analytics: Consolidado de Extracciones")
    with col_header2:
        if st.button("Actualizar Tabla", icon=":material/refresh:", use_container_width=True):
            st.rerun()

    with st.container(border=True):
        if os.path.exists(ARCHIVO_EXCEL):
            try:
                df_extracciones = pd.read_excel(ARCHIVO_EXCEL)
                df_extracciones = df_extracciones.astype(str)
                st.dataframe(df_extracciones, use_container_width=True)

                st.divider()
                with open(ARCHIVO_EXCEL, "rb") as file:
                    st.download_button(
                        label="Descargar Reporte (Excel)",
                        data=file,
                        file_name="reporte_extracciones.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        icon=":material/download:"
                    )
            except Exception as e:
                st.error(f"No se pudo leer el archivo Excel: {e}", icon=":material/warning:")
        else:
            st.info("No existen registros de extracción actuales. Procesa un documento primero.", icon=":material/hourglass_empty:")


# ==========================================
# PESTAÑA 3: GESTIÓN DE VEHÍCULOS (FLUJO UNIFICADO)
# ==========================================
with tab_vehiculos:
    st.subheader(":material/dns: Administración de Base de Datos")

    try:
        df_marcas = pd.read_sql("SELECT * FROM marca ORDER BY nombre ASC", engine)
        df_modelos = pd.read_sql("SELECT * FROM modelo ORDER BY nombre ASC", engine)

        with st.container(border=True):
            st.markdown("### :material/business: Selección de Marca")

            opciones_marcas = df_marcas['id_marca'].tolist() if not df_marcas.empty else []

            if opciones_marcas:
                id_marca_sel = st.selectbox(
                    "Elige una marca del catálogo:",
                    opciones_marcas,
                    format_func=lambda x: df_marcas[df_marcas['id_marca'] == x]['nombre'].values[0]
                )
                marca_actual = df_marcas[df_marcas['id_marca'] == id_marca_sel].iloc[0]
            else:
                id_marca_sel = None
                st.info("No hay marcas en la base de datos. Comienza añadiendo una.", icon=":material/info:")
                marca_actual = None

            col_m1, col_m2, col_m3 = st.columns(3)

            with col_m1:
                with st.expander(":material/add: Añadir Nueva Marca", expanded=False):
                    with st.form("form_add_marca", clear_on_submit=True, border=False):
                        nueva_m_nombre = st.text_input("Nombre de la Marca")
                        nueva_m_origen = st.text_input("Origen (Ej. Japón)")
                        if st.form_submit_button("Guardar Marca", type="primary", use_container_width=True):
                            if nueva_m_nombre:
                                with Session(engine) as session:
                                    session.execute(
                                        text("INSERT INTO marca (nombre, origen) VALUES (:n, :o)"),
                                        {"n": nueva_m_nombre.strip().upper(), "o": nueva_m_origen}
                                    )
                                    session.commit()
                                st.success("Marca guardada.", icon=":material/check:")
                                time.sleep(1)
                                st.rerun()

            if id_marca_sel:
                origen_actual = ""
                if "origen" in marca_actual.index and pd.notna(marca_actual['origen']):
                    origen_actual = str(marca_actual['origen'])

                with col_m2:
                    with st.expander(":material/edit: Editar Seleccionada", expanded=False):
                        with st.form("form_edit_marca", border=False):
                            edit_m_nombre = st.text_input("Nuevo Nombre", value=marca_actual['nombre'])
                            edit_m_origen = st.text_input("Nuevo Origen", value=origen_actual)
                            if st.form_submit_button("Actualizar Marca", use_container_width=True):
                                with Session(engine) as session:
                                    session.execute(
                                        text("UPDATE marca SET nombre = :n, origen = :o WHERE id_marca = :id"),
                                        {"n": edit_m_nombre.strip().upper(), "o": edit_m_origen, "id": id_marca_sel}
                                    )
                                    session.commit()
                                st.success("Marca actualizada.", icon=":material/check:")
                                time.sleep(1)
                                st.rerun()

                with col_m3:
                    with st.expander(":material/delete: Borrar Seleccionada", expanded=False):
                        st.warning(f"⚠️ Se borrará **{marca_actual['nombre']}** y todos sus modelos.")
                        if st.button("Confirmar Eliminación", type="primary", use_container_width=True, key="del_marca"):
                            with Session(engine) as session:
                                session.execute(text("DELETE FROM modelo WHERE id_marca = :id"), {"id": id_marca_sel})
                                session.execute(text("DELETE FROM marca WHERE id_marca = :id"), {"id": id_marca_sel})
                                session.commit()
                            st.error("Marca eliminada.", icon=":material/delete:")
                            time.sleep(1)
                            st.rerun()

        if id_marca_sel:
            st.write("")
            with st.container(border=True):
                st.markdown(f"### :material/directions_car: Modelos de **{marca_actual['nombre']}**")

                modelos_vinculados = df_modelos[df_modelos['id_marca'] == id_marca_sel]

                if not modelos_vinculados.empty:
                    columnas = [c for c in ['id_modelo', 'nombre', 'tipo_carroceria'] if c in modelos_vinculados.columns]
                    vista_modelos = modelos_vinculados[columnas].copy()
                    vista_modelos.columns = ['ID', 'Modelo', 'Carrocería'][:len(columnas)]
                    st.dataframe(vista_modelos.astype(str), use_container_width=True, hide_index=True)
                else:
                    st.info(f"No hay modelos registrados para {marca_actual['nombre']}.", icon=":material/info:")

                opciones_modelos = modelos_vinculados['id_modelo'].tolist() if not modelos_vinculados.empty else []

                if opciones_modelos:
                    id_modelo_sel = st.selectbox(
                        "Elige un modelo de la tabla para gestionarlo:",
                        opciones_modelos,
                        format_func=lambda x: modelos_vinculados[modelos_vinculados['id_modelo'] == x]['nombre'].values[0]
                    )
                    modelo_actual = modelos_vinculados[modelos_vinculados['id_modelo'] == id_modelo_sel].iloc[0]
                else:
                    id_modelo_sel = None

                col_mo1, col_mo2, col_mo3 = st.columns(3)

                with col_mo1:
                    with st.expander(":material/add: Añadir Nuevo Modelo", expanded=False):
                        with st.form("form_add_modelo", clear_on_submit=True, border=False):
                            nuevo_mo_nombre = st.text_input("Nombre del Modelo")
                            nuevo_mo_carroc = st.text_input("Segmento/Carrocería")
                            if st.form_submit_button("Guardar Modelo", type="primary", use_container_width=True):
                                if nuevo_mo_nombre:
                                    with Session(engine) as session:
                                        session.execute(
                                            text("INSERT INTO modelo (id_marca, nombre, tipo_carroceria) VALUES (:m, :n, :t)"),
                                            {"m": id_marca_sel, "n": nuevo_mo_nombre.strip().upper(), "t": nuevo_mo_carroc}
                                        )
                                        session.commit()
                                    st.success("Modelo guardado.", icon=":material/check:")
                                    time.sleep(1)
                                    st.rerun()

                if id_modelo_sel:
                    carroceria_actual = ""
                    if "tipo_carroceria" in modelo_actual.index and pd.notna(modelo_actual['tipo_carroceria']):
                        carroceria_actual = str(modelo_actual['tipo_carroceria'])

                    with col_mo2:
                        with st.expander(":material/edit: Editar Seleccionado", expanded=False):
                            with st.form("form_edit_modelo", border=False):
                                edit_mo_nombre = st.text_input("Nombre", value=modelo_actual['nombre'])
                                edit_mo_carroc = st.text_input("Carrocería", value=carroceria_actual)
                                if st.form_submit_button("Actualizar Modelo", use_container_width=True):
                                    with Session(engine) as session:
                                        session.execute(
                                            text("UPDATE modelo SET nombre = :n, tipo_carroceria = :t WHERE id_modelo = :id"),
                                            {"n": edit_mo_nombre.strip().upper(), "t": edit_mo_carroc, "id": id_modelo_sel}
                                        )
                                        session.commit()
                                    st.success("Modelo actualizado.", icon=":material/check:")
                                    time.sleep(1)
                                    st.rerun()

                    with col_mo3:
                        with st.expander(":material/delete: Borrar Seleccionado", expanded=False):
                            st.warning(f"⚠️ Se borrará el modelo **{modelo_actual['nombre']}**.")
                            if st.button("Confirmar Eliminación", type="primary", use_container_width=True, key="del_modelo"):
                                with Session(engine) as session:
                                    session.execute(text("DELETE FROM modelo WHERE id_modelo = :id"), {"id": id_modelo_sel})
                                    session.commit()
                                st.error("Modelo eliminado.", icon=":material/delete:")
                                time.sleep(1)
                                st.rerun()

    except Exception as e:
        st.error(f"Fallo de conexión con la base de datos: {e}", icon=":material/cloud_off:")
