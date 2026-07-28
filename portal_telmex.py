import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import glob
import datetime
import requests
import zipfile
import io
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Tablero Operativo Bolsas", layout="wide")

# ---------------------------------------------------------
# 🎨 ESTILOS CORPORATIVOS (OCULTAR ICONOS DE STREAMLIT/GITHUB)
# ---------------------------------------------------------
ocultar_iconos = """
<style>
/* 1. Ocultar el encabezado completo (desaparece Fork, GitHub y Menú) */
[data-testid="stHeader"] {
    display: none !important;
}

/* 2. Ocultar barra de herramientas secundaria por seguridad */
[data-testid="stToolbar"] {
    display: none !important;
}

/* 3. Ocultar el menú de hamburguesa nativo */
#MainMenu {
    display: none !important;
}

/* 4. Ocultar pie de página (marca de agua de Streamlit) */
footer {
    display: none !important;
}

/* 5. Ocultar el espacio en blanco que deja el encabezado al desaparecer */
.stApp > header {
    background-color: transparent !important;
}
</style>
"""
st.markdown(ocultar_iconos, unsafe_allow_html=True)
# ---------------------------------------------------------

# ---------------------------------------------------------
# 🎛️ SELECTOR DE REGIÓN (NUEVO)
# ---------------------------------------------------------
# Centramos el título y el selector para que se vea elegante
st.markdown("<h1 style='text-align: center;'>📊 Tablero Operativo - Bolsas</h1>", unsafe_allow_html=True)

col_espacio1, col_selector, col_espacio2 = st.columns([1, 2, 1])
with col_selector:
    region_seleccionada = st.selectbox(
        "📍 Selecciona la Región a visualizar:",
        ["Monterrey", "Tamaulipas"]
    )
# ---------------------------------------------------------

# --- INYECCIÓN DE CSS PARA COMPACTAR, ESCALAR Y HACER RESPONSIVAS LAS TABLAS ---
st.markdown("""
    <style>
        /* 💻 ESTILOS GENERALES */
        [data-testid="stTable"] { 
            width: max-content !important; 
            max-width: 100%; 
            overflow-x: auto; 
            font-size: 85% !important; 
        }
        [data-testid="stTable"] table { width: auto !important; }
        
        /* 👇 SOLO centrar los encabezados superiores (Columnas) */
        [data-testid="stTable"] thead th {
            text-align: center !important;
            vertical-align: middle !important;
        }

        /* 👇 Mantener a la izquierda la columna de CTs (Índices) */
        [data-testid="stTable"] tbody th {
            text-align: left !important;
        }

        [data-testid="stTable"] th, [data-testid="stTable"] td {
            white-space: nowrap !important;
            padding: 5px 10px !important; 
        }

        /* 📱 ESTILOS PARA CELULARES */
        @media (max-width: 768px) {
            [data-testid="stTable"] th, [data-testid="stTable"] td {
                font-size: 10px !important; 
                padding: 4px 6px !important; 
            }
        }
    </style>
""", unsafe_allow_html=True)
st.divider()
# ---------------------------------------------------------
# ☁️ LÓGICA DE DETECCIÓN AUTOMÁTICA (CLARO DRIVE)
# ---------------------------------------------------------
def obtener_archivo_clarodrive():
    # Truco de ClaroDrive: Agregamos /download a tu liga para bajar la carpeta
    url_carpeta = "https://i0000.clarodrive.com/s/KRrAxbKcriJiwcK"
    url_descarga = url_carpeta.rstrip('/') + '/download'
    
    try:
        respuesta = requests.get(url_descarga, timeout=15)
        if respuesta.status_code == 200:
            # Leemos el archivo ZIP directamente en la memoria del servidor
            with zipfile.ZipFile(io.BytesIO(respuesta.content)) as archivo_zip:
                # Buscamos todos los archivos Excel (ignorando los temporales que empiezan con ~)
                excel_infos = [info for info in archivo_zip.infolist() if info.filename.endswith('.xlsx') and not info.filename.startswith('~')]
                
                if excel_infos:
                    # Si hay varios, tomamos el más reciente por fecha de modificación
                    excel_reciente = max(excel_infos, key=lambda x: x.date_time)
                    
                    # Lo extraemos a la memoria
                    archivo_bytes = io.BytesIO(archivo_zip.read(excel_reciente.filename))
                    
                    # =========================================================
                    # 🕒 AJUSTE DE ZONA HORARIA (UTC A CENTRO DE MÉXICO)
                    # =========================================================
                    fecha_tupla = excel_reciente.date_time 
                    
                    # 1. Convertimos la tupla del ZIP a un formato de fecha manipulable
                    fecha_utc = datetime.datetime(
                        year=fecha_tupla[0], month=fecha_tupla[1], day=fecha_tupla[2],
                        hour=fecha_tupla[3], minute=fecha_tupla[4], second=fecha_tupla[5]
                    )
                    
                    # 2. Le restamos 6 horas (Diferencia de México respecto a UTC)
                    fecha_mexico = fecha_utc - datetime.timedelta(hours=6)
                    
                    # 3. Lo convertimos al texto final
                    fecha_str = fecha_mexico.strftime('%d/%m/%Y %H:%M:%S')
                    # =========================================================
                    
                    return archivo_bytes, excel_reciente.filename, fecha_str
    except Exception:
        pass # Si falla el internet del servidor o la liga, no rompe el programa
    
    return None, None, None

archivo_automatico, nombre_corto, fecha_actualizacion = obtener_archivo_clarodrive()
archivo_a_procesar = None

col1, col2 = st.columns([2, 1])
with col1:
    if archivo_automatico:
        st.success(f"☁️ **Base de datos (Sharepoint):** {nombre_corto}  \n⏱️ **Actualizado:** {fecha_actualizacion}")
        archivo_a_procesar = archivo_automatico
    else:
        st.warning("⚠️ No se pudo conectar con Claro Drive o la carpeta está vacía.")

with col2:
    # Si Claro Drive falla, habilitamos la subida manual como "Plan B"
    usar_manual = st.checkbox("Subir archivo manualmente", value=False if archivo_automatico else True)

if usar_manual:
    archivo_a_procesar = st.file_uploader("Arrastra aquí tu archivo de Excel", type=['xlsx'])

st.divider()

# ---------------------------------------------------------
# PROCESAMIENTO
# ---------------------------------------------------------
if archivo_a_procesar is not None:
    datos_listos = False
    with st.spinner(f"⏳ Estructurando datos para {region_seleccionada}..."):
        try:
            df = pd.read_excel(archivo_a_procesar, sheet_name='Detalle')
            df.columns = df.columns.str.strip()
            
# Limpieza Básica
            columnas_clave = ['ULTIMOS_6_MESES', 'ESTATUS_AGR_N2', 'ESTATUS_AGR_N1', 'ETAPA_OS', 'PORTABILIDAD', 'TIPO_MOVIMIENTO', 'ESTATUS_OS']
            for col in columnas_clave:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()

            # 🧹 1. FILTRO DE ESTATUS: Conservar solo PENDIENTES
            if 'ESTATUS_OS' in df.columns:
                df = df[df['ESTATUS_OS'] == 'PENDIENTE']

            # 🧹 2. FILTRO DE CT Y "PASES VIP" PARA BLANCOS
            if 'CT' in df.columns:
                # Adiós Tulancingo incondicionalmente
                df = df[df['CT'].astype(str).str.upper() != 'CT TULANCINGO']
                
                # Identificamos cuáles CT vienen vacíos o nulos
                es_ct_vacio = df['CT'].isna() | (df['CT'].astype(str).str.strip() == '')
                
                # Identificamos los estatus que tienen permitido no tener CT
                estatus_n2 = df['ESTATUS_AGR_N2'].astype(str).str.upper()
                pase_vip = estatus_n2.str.contains('4 EN VALIDACION') | estatus_n2.str.contains('6. PENDIENTE')
                
                # REGLA: Conservar los que SÍ tienen CT, o los que NO tienen pero tienen PASE VIP
                df = df[~es_ct_vacio | (es_ct_vacio & pase_vip)]
                
                # Rellenamos los vacíos permitidos con "SIN CT" para que Pandas no choque en otras tablas
                df['CT'] = df['CT'].fillna('SIN CT')
                df.loc[df['CT'].astype(str).str.strip() == '', 'CT'] = 'SIN CT'

            # =========================================================
            # 🗺️ LÓGICA DE RUTEO POR REGIÓN
            # =========================================================
            if region_seleccionada == "Monterrey":
                diccionario_ct_area = {
                    "CT SANTA FE [MTY]": "MONTERREY 2", "CT UNIVERSIDAD (NL)": "MONTERREY 2", "CT LA SILLA": "MONTERREY 2",
                    "CT PUENTES": "MONTERREY 2", "CT SANTA CATARINA": "MONTERREY 2",
                    "CT REVOLUCION": "MONTERREY 1", "CT LINCOLN": "MONTERREY 1", "CT SAN PEDRO [NL]": "MONTERREY 1",
                    "CT COLON (MTY)": "MONTERREY 1", "CT GONZALITOS": "MONTERREY 1",
                    "CT GENERAL ESCOBEDO (BRISAS)": "MONTERREY 3", "CT CADEREYTA": "MONTERREY 3",
                    "CT MONTEMORELOS": "MONTERREY 3", "CT APODACA": "MONTERREY 3", "CT LINARES": "MONTERREY 3"
                }
                df['AREA_CORREGIDA'] = df['CT'].map(diccionario_ct_area).fillna(df.get('AREA', "ÁREA NO DEFINIDA"))
                df = df[df['AREA_CORREGIDA'].isin(["MONTERREY 1", "MONTERREY 2", "MONTERREY 3"])]
                
            elif region_seleccionada == "Tamaulipas":
                # Asumimos que la columna 'AREA' tiene la info. Lo pasamos a mayúsculas.
                df['AREA_TEMP'] = df.get('AREA', '').astype(str).str.strip().str.upper()
                
                # Agrupamos Matamoros y Reynosa, y normalizamos variaciones de Ciudad Victoria
                diccionario_tam = {
                    "MATAMOROS": "MATAMOROS-REYNOSA",
                    "REYNOSA": "MATAMOROS-REYNOSA",
                    "CD VICTORIA": "CIUDAD VICTORIA",
                    "CD. VICTORIA": "CIUDAD VICTORIA",
                    "VICTORIA": "CIUDAD VICTORIA"
                }
                df['AREA_CORREGIDA'] = df['AREA_TEMP'].map(diccionario_tam).fillna(df['AREA_TEMP'])
                df = df[df['AREA_CORREGIDA'].isin(["CIUDAD VICTORIA", "NUEVO LAREDO", "MATAMOROS-REYNOSA", "TAMPICO"])]
            # =========================================================

            # Rangos
            if 'DIL2' in df.columns:
                df['DIL2'] = pd.to_numeric(df['DIL2'], errors='coerce')
                cortes = [-float('inf'), 2, 5, 10, 20, 50, 100, float('inf')]
                etiquetas = ['0 A 2', '3 A 5', '6 A 10', '11 A 20', '21 A 50', '51 A 100', '> 100']
                df['Rango x Dil'] = pd.cut(df['DIL2'], bins=cortes, labels=etiquetas).astype(str).replace('nan', 'Sin Rango')
            
            datos_listos = True
        except Exception as e:
            st.error(f"Error al procesar los datos: {e}")
            st.stop()

    if datos_listos:
        orden_columnas = ['0 A 2', '3 A 5', '6 A 10', '11 A 20', '21 A 50', '51 A 100', '> 100', 'Sin Rango', 'Total']

        # -----------------------------------------------------
        # 🛠️ FUNCIONES DE FORMATO Y LÓGICA
        # -----------------------------------------------------
        def ordenar_numerico(texto):
            if texto == 'Total': return 9999
            match = re.search(r'^(\d+)', str(texto))
            return int(match.group(1)) if match else 999

        def aplicar_subtotales(pt):
            if not isinstance(pt.index, pd.MultiIndex) or pt.empty: return pt
            if 'Total' in pt.index.get_level_values(0):
                pt_sin_total = pt.drop('Total', level=0)
                total_gen = pt.loc[['Total']]
            else:
                pt_sin_total = pt
                total_gen = pd.DataFrame()
                
            subtotales = pt_sin_total.groupby(level=0).sum()
            subtotales.index = pd.MultiIndex.from_arrays([subtotales.index, ['ZZZ_SUBTOTAL'] * len(subtotales.index)])
            pt_completa = pd.concat([pt_sin_total, subtotales]).sort_index()
            pt_completa = pt_completa.rename(index={'ZZZ_SUBTOTAL': '👉 TOTAL ÁREA'})
            
            if not total_gen.empty: 
                pt_completa = pd.concat([pt_completa, total_gen])
            return pt_completa

        def estilo_resaltado(df):
            estilo = df.style.apply(
                lambda x: ['background-color: #ADD8E6; font-weight: bold' if isinstance(x.name, tuple) and len(x.name) > 1 and x.name[1] == '👉 TOTAL ÁREA' else '' for _ in x], 
                axis=1
            )
            return estilo

        # 📥 FUNCIÓN PARA BOTONES DE DESCARGA DINÁMICOS
        def generar_boton_descarga(df_datos, base_nombre, btn_key):
            # Le agregamos la región al nombre del archivo descargado
            nombre_final = f"{base_nombre}_{region_seleccionada.lower()}.csv"
            csv = df_datos.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📥 Descargar Folios",
                data=csv,
                file_name=nombre_final,
                mime='text/csv',
                key=btn_key
            )

        def estilo_tabla_1(df):
            def aplicar_colores(row):
                nombre = str(row.name).replace('\u200b', '')
                nombre_limpio = nombre.strip().upper()
                es_subtotal_n2 = not nombre.startswith('\xa0') and nombre_limpio != 'TOTAL GENERAL'
                
                if es_subtotal_n2 and '4 EN VALIDACION' in nombre_limpio:
                    return ['background-color: #FF0000; color: white; font-weight: bold'] * len(row)
                elif es_subtotal_n2 and '9 VENTAS' in nombre_limpio:
                    return ['background-color: #F4A460; color: black; font-weight: bold'] * len(row)
                elif '6.3 EN PROCESO' in nombre_limpio:
                    return ['background-color: #FFFF00; color: black; font-weight: bold'] * len(row)
                elif '6.4 MANTENIMIENTO' in nombre_limpio:
                    return ['background-color: #FFE4B5; color: black'] * len(row)
                elif es_subtotal_n2:
                    return ['background-color: #ADD8E6; color: black; font-weight: bold'] * len(row)
                elif nombre_limpio == 'TOTAL GENERAL':
                    return ['font-weight: bold; background-color: #EFEFEF'] * len(row)
                return [''] * len(row)
            return df.style.apply(aplicar_colores, axis=1)

        # -----------------------------------------------------
        # 📊 DIBUJADO DE TABLAS Y BOTONES
        # -----------------------------------------------------
        
        # 1. Resumen General
        st.subheader(f"📑 1. Últimos 6 Meses (Demanda por Bolsa) - {region_seleccionada}")
        df_6m = df[df['ULTIMOS_6_MESES'] == 'SI']
        estatus_excluidos = ["7. ASPECTOS TÉCNICOS", "8 PENDIENTES BOLSA VENTAS + CD", "10 NO VENTAS + CD"]
        df_6m = df_6m[~df_6m['ESTATUS_AGR_N2'].isin(estatus_excluidos)]

        if not df_6m.empty:
            td_6m = pd.pivot_table(df_6m, index=['ESTATUS_AGR_N2', 'ESTATUS_AGR_N1'], columns='AREA_CORREGIDA', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            
            total_gen_6m = td_6m.loc[['Total']] if 'Total' in td_6m.index.get_level_values(0) else pd.DataFrame()
            td_6m_sin_total = td_6m.drop('Total', level=0) if 'Total' in td_6m.index.get_level_values(0) else td_6m
            
            subtotales_n2 = td_6m_sin_total.groupby(level=0).sum()
            subtotales_n2.index = pd.MultiIndex.from_arrays([subtotales_n2.index, [''] * len(subtotales_n2.index)])
            
            td_6m_final = pd.concat([td_6m_sin_total, subtotales_n2])
            nuevo_indice = sorted(td_6m_final.index, key=lambda x: (ordenar_numerico(x[0]), str(x[1])))
            td_6m_final = td_6m_final.reindex(nuevo_indice)
            
            if not total_gen_6m.empty:
                td_6m_final = pd.concat([td_6m_final, total_gen_6m])
            
            td_6m_final.index.names = ['ESTATUS_AGR_N2', 'ESTATUS_AGR_N1']
            td_6m_final = td_6m_final.reset_index()
            
            vistos = {}
            def unificar_estatus(fila):
                n2 = str(fila['ESTATUS_AGR_N2']) if pd.notna(fila['ESTATUS_AGR_N2']) else ''
                n1 = str(fila['ESTATUS_AGR_N1']) if pd.notna(fila['ESTATUS_AGR_N1']) else ''
                if n2 == 'Total': texto = 'Total General'
                elif n1 == '': texto = n2  
                else: texto = '\xa0\xa0\xa0\xa0\xa0\xa0' + n1
                if texto in vistos: vistos[texto] += 1
                else: vistos[texto] = 0
                return texto + ('\u200b' * vistos[texto])
            
            td_6m_final['ESTATUS_AGR'] = td_6m_final.apply(unificar_estatus, axis=1)
            td_6m_final = td_6m_final.set_index('ESTATUS_AGR')
            td_6m_final = td_6m_final.drop(columns=['ESTATUS_AGR_N2', 'ESTATUS_AGR_N1'])
            
            st.table(estilo_tabla_1(td_6m_final))
            generar_boton_descarga(df_6m, 'folios_t1_6m', btn_key='btn1')
        else:
            st.info("No hay datos para esta tabla.")
        
        st.divider()

        # 2. Bolsa 4
        st.subheader("✔ 2. Ult 6 Meses (BOLSA 4xDIL2)")
        df_b4 = df[df['ESTATUS_AGR_N2'].str.contains('4', case=False, na=False)]
        if not df_b4.empty:
            td_b4 = pd.pivot_table(df_b4, index=['AREA_CORREGIDA', 'ESTATUS_AGR_N1'], columns='Rango x Dil', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            cols = [c for c in orden_columnas if c in td_b4.columns] + [c for c in td_b4.columns if c not in orden_columnas]
            st.table(estilo_resaltado(aplicar_subtotales(td_b4[cols])))
            generar_boton_descarga(df_b4, 'folios_t2_bolsa4', btn_key='btn2')
        else: st.info("No hay datos.")

        st.divider()

        # 3. Bolsa 5xDIL2 - Solo PS
        st.subheader("🛠 3. Ult 6 Meses (BOLSA 5xDIL2 - Solo PS)")
        
        df_b5_ps = df[(df['ESTATUS_AGR_N2'].str.contains('5', case=False, na=False)) & (df['ETAPA_OS'] == 'PS')]
        if not df_b5_ps.empty:
            
            # 🔘 NUEVO: Interruptor para ocultar/mostrar la captura
            mostrar_captura = st.toggle("📝 Compromiso Produccion", value=False)
            
            # Procesamos tabla izquierda (Esto se hace siempre para mostrar los datos base)
            td_b5_ps = pd.pivot_table(df_b5_ps, index=['AREA_CORREGIDA', 'CT'], columns='Rango x Dil', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            cols = [c for c in orden_columnas if c in td_b5_ps.columns] + [c for c in td_b5_ps.columns if c not in orden_columnas]
            df_con_subtotales = aplicar_subtotales(td_b5_ps[cols])
            
            # Lógica de columnas dinámicas
            if mostrar_captura:
                # Si está prendido, partimos la pantalla en dos
                col_t3_izq, col_t3_der = st.columns([1.1, 1.1]) 
            else:
                # Si está apagado, la tabla izquierda usa todo su espacio natural
                col_t3_izq = st.container()
            
            with col_t3_izq:
                st.table(estilo_resaltado(df_con_subtotales))
                generar_boton_descarga(df_b5_ps, 'folios_t3_bolsa5_ps', btn_key='btn3')
            
            # ⬇️ Todo el bloque de Google Sheets se ejecuta SOLO si el switch está prendido
            if mostrar_captura:
                with col_t3_der:
                    st.markdown("<div style='height: 38px;'></div>", unsafe_allow_html=True)
                    
                    # 1. 🌐 CONECTAR A GOOGLE SHEETS
                    url_sheet = "https://docs.google.com/spreadsheets/d/1cYHq6afeVavGNGSRPC3lPxzLN5E5wIqFFa2BSp0nxZY/edit?gid=0#gid=0"
                    try:
                        conn = st.connection("gsheets", type=GSheetsConnection)
                        df_historico = conn.read(spreadsheet=url_sheet, worksheet=region_seleccionada, ttl=0, usecols=[0,1,2,3,4,5])
                    except Exception as e:
                        df_historico = pd.DataFrame()
                    
                    # 2. BASE DE CTS
                    df_captura = df_con_subtotales.copy().reset_index()
                    df_base = pd.DataFrame({
                        'AREA': df_captura['AREA_CORREGIDA'] if 'AREA_CORREGIDA' in df_captura.columns else df_captura.iloc[:, 0],
                        'CT': df_captura['CT'] if 'CT' in df_captura.columns else df_captura.iloc[:, 1],
                        'Folios': df_captura['Total'] if 'Total' in df_captura.columns else 0
                    })
                    df_base.loc[df_base['AREA'].astype(str).str.lower() == 'total', 'CT'] = 'GRAN TOTAL'
                    
                    # 3. MERGE CORREGIDO
                    if not df_historico.empty and 'CT' in df_historico.columns:
                        df_hist_limpio = df_historico[~df_historico['CT'].astype(str).str.contains('TOTAL', case=False, na=False)].drop_duplicates(subset=['CT'])
                        df_merge = pd.merge(df_base, df_hist_limpio[['CT', 'SIAC', 'Tecs', 'Comp', 'Prod. Esp']], on='CT', how='left')
                    else:
                        df_merge = df_base.copy()
                        df_merge['SIAC'] = 0
                        df_merge['Tecs'] = 0
                        df_merge['Comp'] = 0.0
                        df_merge['Prod. Esp'] = 0
                    
                    df_merge['SIAC'] = df_merge['SIAC'].fillna(0).astype(int)
                    df_merge['Tecs'] = df_merge['Tecs'].fillna(0).astype(int)
                    df_merge['Comp'] = df_merge['Comp'].fillna(0.0)
                    df_merge['Prod. Esp'] = df_merge['Prod. Esp'].fillna(0).astype(int)
                    
                    # 🔄 AUTO-SUMAS INICIALES
                    mask_cts_ini = (~df_merge['CT'].astype(str).str.contains('TOTAL', case=False)) & (df_merge['AREA'].astype(str).str.lower() != 'total')
                    for area in df_merge['AREA'].unique():
                        if str(area).lower() != 'total':
                            mask_area_cts_ini = mask_cts_ini & (df_merge['AREA'] == area)
                            mask_sub_ini = (df_merge['AREA'] == area) & (df_merge['CT'].astype(str).str.contains('TOTAL', case=False))
                            if mask_sub_ini.any():
                                df_merge.loc[mask_sub_ini, 'SIAC'] = df_merge.loc[mask_area_cts_ini, 'SIAC'].sum()
                                df_merge.loc[mask_sub_ini, 'Tecs'] = df_merge.loc[mask_area_cts_ini, 'Tecs'].sum()
                                df_merge.loc[mask_sub_ini, 'Prod. Esp'] = df_merge.loc[mask_area_cts_ini, 'Prod. Esp'].sum()
                                df_merge.loc[mask_sub_ini, 'Comp'] = 0
                                
                    mask_gran_ini = df_merge['AREA'].astype(str).str.lower() == 'total'
                    if mask_gran_ini.any():
                        df_merge.loc[mask_gran_ini, 'SIAC'] = df_merge.loc[mask_cts_ini, 'SIAC'].sum()
                        df_merge.loc[mask_gran_ini, 'Tecs'] = df_merge.loc[mask_cts_ini, 'Tecs'].sum()
                        df_merge.loc[mask_gran_ini, 'Prod. Esp'] = df_merge.loc[mask_cts_ini, 'Prod. Esp'].sum()
                        df_merge.loc[mask_gran_ini, 'Comp'] = 0

                    df_mostrar = df_merge.drop(columns=['AREA'])
                    
                    def pintar_totales(row):
                        if 'TOTAL ÁREA' in str(row['CT']):
                            return ['background-color: #ADD8E6; font-weight: bold; color: black;'] * len(row)
                        return [''] * len(row)
                    
                    df_estilizado = df_mostrar.style.apply(pintar_totales, axis=1)
                    alto_dinamico = int((len(df_mostrar) * 36) + 78) 
                    
                    # 4. MOSTRAR EDITOR
                    df_editado = st.data_editor(
                        df_estilizado, 
                        hide_index=True,
                        use_container_width=True,
                        height=alto_dinamico,
                        disabled=['CT', 'Folios', 'Prod. Esp'], 
                        column_config={
                            "CT": st.column_config.TextColumn("CT\n", width="medium"),
                            "Folios": st.column_config.NumberColumn("Folios\n", width=50),
                            "SIAC": st.column_config.NumberColumn("SIAC\n", width=60),
                            "Tecs": st.column_config.NumberColumn("Tecs\n", width=60),
                            "Comp": st.column_config.NumberColumn("Comp\n", width=80),
                            "Prod. Esp": st.column_config.NumberColumn("Prod. Esp\n", width=80, format="%d")
                        }
                    )
                    
                    # 5. BOTÓN MAESTRO: Calcula y Direcciona al Acumulado Correcto
                    if st.button(f"☁️ Guardar y Acumular ({region_seleccionada})", type="primary"):
                        df_guardar = df_editado.copy()
                        
                        df_guardar['Prod. Esp'] = (df_guardar['Tecs'] * df_guardar['Comp']).round().astype(int)
                        mask_cts = (~df_base['CT'].astype(str).str.contains('TOTAL', case=False)) & (df_base['AREA'].astype(str).str.lower() != 'total')
                        
                        for area in df_base['AREA'].unique():
                            if str(area).lower() != 'total':
                                mask_area_cts = mask_cts & (df_base['AREA'] == area)
                                mask_sub = (df_base['AREA'] == area) & (df_base['CT'].astype(str).str.contains('TOTAL', case=False))
                                if mask_sub.any():
                                    df_guardar.loc[mask_sub, 'SIAC'] = df_guardar.loc[mask_area_cts, 'SIAC'].sum()
                                    df_guardar.loc[mask_sub, 'Tecs'] = df_guardar.loc[mask_area_cts, 'Tecs'].sum()
                                    df_guardar.loc[mask_sub, 'Prod. Esp'] = df_guardar.loc[mask_area_cts, 'Prod. Esp'].sum()
                                    df_guardar.loc[mask_sub, 'Comp'] = 0
                                    
                        mask_gran = df_base['AREA'].astype(str).str.lower() == 'total'
                        if mask_gran.any():
                            df_guardar.loc[mask_gran, 'SIAC'] = df_guardar.loc[mask_cts, 'SIAC'].sum()
                            df_guardar.loc[mask_gran, 'Tecs'] = df_guardar.loc[mask_cts, 'Tecs'].sum()
                            df_guardar.loc[mask_gran, 'Prod. Esp'] = df_guardar.loc[mask_cts, 'Prod. Esp'].sum()
                            df_guardar.loc[mask_gran, 'Comp'] = 0
                        
                        df_final_region = df_guardar[['CT', 'Folios', 'SIAC', 'Tecs', 'Comp', 'Prod. Esp']]
                        conn.update(spreadsheet=url_sheet, worksheet=region_seleccionada, data=df_final_region)
                        
                        # C. LÓGICA DEL ACUMULADO DINÁMICO
                        hoja_acumulado = "AcumuladoM" if "Monterrey" in region_seleccionada else "AcumuladoT"
                        fecha_hoy = datetime.datetime.now().strftime("%d/%m/%Y")
                        
                        df_hist_nuevo = df_final_region.copy()
                        df_hist_nuevo.insert(0, 'Fecha', fecha_hoy)
                        df_hist_nuevo.insert(1, 'Region', region_seleccionada)
                        
                        try:
                            df_acumulado_base = conn.read(spreadsheet=url_sheet, worksheet=hoja_acumulado, ttl=0)
                            if not df_acumulado_base.empty and 'Fecha' in df_acumulado_base.columns:
                                mask_duplicados = (df_acumulado_base['Fecha'] == fecha_hoy) & (df_acumulado_base['Region'] == region_seleccionada)
                                df_acumulado_limpio = df_acumulado_base[~mask_duplicados]
                            else:
                                df_acumulado_limpio = pd.DataFrame()
                        except:
                            df_acumulado_limpio = pd.DataFrame()
                        
                        df_final_acumulado = pd.concat([df_acumulado_limpio, df_hist_nuevo], ignore_index=True)
                        conn.update(spreadsheet=url_sheet, worksheet=hoja_acumulado, data=df_final_acumulado)

                        st.success(f"¡Datos de {region_seleccionada} guardados exitosamente!")
                        st.rerun() 
        else: 
            st.info("No hay datos para la Bolsa 5 - Solo PS.")
        
        st.divider()
        
        # 4. Bolsa 5 por Etapa
        st.subheader("⚙ 4. Ult 6 Meses (BOLSA 5 por Etapa)")
        df_b5 = df[df['ESTATUS_AGR_N2'].str.contains('5', case=False, na=False)]
        if not df_b5.empty:
            td_b5 = pd.pivot_table(df_b5, index=['AREA_CORREGIDA', 'CT'], columns='ETAPA_OS', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            st.table(estilo_resaltado(aplicar_subtotales(td_b5)))
            generar_boton_descarga(df_b5, 'folios_t4_bolsa5_etapas', btn_key='btn4')
        else: st.info("No hay datos.")

        st.divider()

        # 5. Portas
        st.subheader("📞 5. Ult 6 Meses (PORTAS)")
        df_p = df[(df['PORTABILIDAD']=='SI') & (df['ULTIMOS_6_MESES']=='SI') & (df['ETAPA_OS']=='PS')]
        if not df_p.empty:
            td_p = pd.pivot_table(df_p, index=['AREA_CORREGIDA', 'CT'], columns='Rango x Dil', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            cols = [c for c in orden_columnas if c in td_p.columns] + [c for c in td_p.columns if c not in orden_columnas]
            st.table(estilo_resaltado(aplicar_subtotales(td_p[cols])))
            generar_boton_descarga(df_p, 'folios_t5_portas', btn_key='btn5')
        else: st.info("No hay datos.")

        st.divider()

        # 6. Cambio de Domicilio
        st.subheader("🏠 6. Ult 6 Meses (CAMB DOM)")        
        df_cd = df[df['TIPO_MOVIMIENTO'].str.contains('CAMB DOM', case=False, na=False)]
        if not df_cd.empty:
            td_cd = pd.pivot_table(df_cd, index=['AREA_CORREGIDA', 'CT'], columns='ETAPA_OS', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            st.table(estilo_resaltado(aplicar_subtotales(td_cd)))
            generar_boton_descarga(df_cd, 'folios_t6_cambio_domicilio', btn_key='btn6')
        else: st.info("No hay datos.")
