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

st.set_page_config(page_title="Tablero Operativo Bolsas", layout="wide")

# ---------------------------------------------------------
# 🎨 ESTILOS CORPORATIVOS (OCULTAR ICONOS DE STREAMLIT/GITHUB)
# ---------------------------------------------------------
ocultar_iconos = """
<style>
/* Oculta el menú de hamburguesa y el icono de GitHub en la esquina superior derecha */
#MainMenu {visibility: hidden;}
[data-testid="stToolbar"] {visibility: hidden;}

/* Oculta el botón de "Deploy" si llega a aparecer */
.stDeployButton {display:none;}

/* Oculta la marca de agua de "Made with Streamlit" en la parte inferior */
footer {visibility: hidden;}
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

# --- INYECCIÓN DE CSS PARA COMPACTAR LAS TABLAS ---
st.markdown("""
    <style>
        [data-testid="stTable"] { width: max-content !important; }
        [data-testid="stTable"] table { width: auto !important; }
        [data-testid="stTable"] th, [data-testid="stTable"] td {
            white-space: nowrap !important;
            padding: 8px 15px !important;
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
                    
                    # Le damos formato a la fecha interna del ZIP
                    fecha_tupla = excel_reciente.date_time 
                    fecha_str = f"{fecha_tupla[2]:02d}/{fecha_tupla[1]:02d}/{fecha_tupla[0]} {fecha_tupla[3]:02d}:{fecha_tupla[4]:02d}:{fecha_tupla[5]:02d}"
                    
                    return archivo_bytes, excel_reciente.filename, fecha_str
    except Exception:
        pass # Si falla el internet del servidor o la liga, no rompe el programa
    
    return None, None, None

archivo_automatico, nombre_corto, fecha_actualizacion = obtener_archivo_clarodrive()
archivo_a_procesar = None

col1, col2 = st.columns([2, 1])
with col1:
    if archivo_automatico:
        st.success(f"☁️ **Base de datos (Claro Drive):** {nombre_corto}  \n⏱️ **Actualizado:** {fecha_actualizacion}")
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
            td_b5_ps = pd.pivot_table(df_b5_ps, index=['AREA_CORREGIDA', 'CT'], columns='Rango x Dil', values='FOLIO', aggfunc='count', fill_value=0, margins=True, margins_name='Total')
            cols = [c for c in orden_columnas if c in td_b5_ps.columns] + [c for c in td_b5_ps.columns if c not in orden_columnas]
            st.table(estilo_resaltado(aplicar_subtotales(td_b5_ps[cols])))
            generar_boton_descarga(df_b5_ps, 'folios_t3_bolsa5_ps', btn_key='btn3')
        else: st.info("No hay datos.")

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
