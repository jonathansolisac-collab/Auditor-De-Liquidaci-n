import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Syntix - Control de Liquidación", layout="wide")

st.title("📊Desarrollo de Jonathan Solis Auditor de liquidación")
st.markdown("Módulo automatizado de validación")

# ==========================================
# 0. CONTROL DE BORRADO DE TEXTO (SESSION STATE)
# ==========================================
# Inicializamos la variable en memoria si no existe
if "txt_remisiones" not in st.session_state:
    st.session_state["txt_remisiones"] = ""

# Función que se ejecuta al presionar el botón de limpiar
def limpiar_caja_texto():
    st.session_state["txt_remisiones"] = ""

# ==========================================
# 1. INTERFAZ DE CARGA Y ENTRADA DE DATOS
# ==========================================
col_files1, col_files2 = st.columns(2)

with col_files1:
    archivo_soriana = st.file_uploader("1. Cargar Base Soriana Portal (Excel)", type=["xlsx"])
    
    st.markdown("##### 2. Pegar Remisiones por Ruta (SICAV):")
    
    # El cuadro de texto ahora se vincula al st.session_state para poder controlarlo
    datos_pegados_remisiones = st.text_area(
        "Copia las celdas de tu reporte (con encabezados) y pégalas aquí (Ctrl+V):", 
        height=150,
        placeholder="Cedis\tRuta\tRemision\tProducto\tPrecio\tCantidad\tSubTotal\n...",
        key="txt_remisiones"
    )
    
    # Botón operativo para limpiar el texto pegado de golpe
    st.button("🗑️ Limpiar datos pegados", on_click=limpiar_caja_texto, use_container_width=True)

with col_files2:
    archivo_relacion = st.file_uploader("3. Cargar Relación de Remisiones (Excel)", type=["xlsx", "xls"])
    archivo_upc = st.file_uploader("4. Cargar Catálogo UPC (Excel)", type=["xlsx"])

# Validar que tengamos todas las entradas de datos arriba
tiene_remisiones = datos_pegados_remisiones.strip() != ""

if archivo_soriana and tiene_remisiones and archivo_relacion and archivo_upc:
    try:
        # ==========================================
        # 2. PROCESAMIENTO Y LIMPIEZA DE BASES
        # ==========================================
        
        # --- BASE SORIANA ---
        df_sor = pd.read_excel(archivo_soriana)
        df_sor.columns = df_sor.columns.astype(str).str.strip()
        
        columnas_a_rellenar = ['Folio de entrada', 'Proveedor', 'Fecha de recibo', 'Sucursal', 'Factura', 'Folio de recibo', 'Total', 'Moneda']
        columnas_presentes = [col for col in columnas_a_rellenar if col in df_sor.columns]
        if columnas_presentes:
            df_sor[columnas_presentes] = df_sor[columnas_presentes].ffill()
        
        if 'Folio de recibo' not in df_sor.columns or 'Código de barras' not in df_sor.columns or 'Cantidad Total' not in df_sor.columns:
            st.error("❌ La base de Soriana debe contener las columnas 'Folio de recibo', 'Código de barras' y 'Cantidad Total'.")
            st.stop()
            
        df_sor['Folio de recibo'] = df_sor['Folio de recibo'].astype(str).str.split('.').str[0].str.strip()
        df_sor['Código de barras'] = df_sor['Código de barras'].astype(str).str.split('.').str[0].str.strip()
        df_sor['Sucursal'] = df_sor['Sucursal'].fillna('').astype(str).str.split('.').str[0].str.strip()
        
        df_sor['Cantidad Total'] = pd.to_numeric(df_sor['Cantidad Total'], errors='coerce').fillna(0)
        
        df_sor['Llave_Soriana'] = df_sor['Folio de recibo'] + df_sor['Código de barras']
        
        df_sor_clean = df_sor.groupby('Llave_Soriana').agg({
            'Folio de recibo': 'first',
            'Código de barras': 'first',
            'Sucursal': 'first',
            'Cantidad Total': 'sum',
            'Descripción': 'first'
        }).reset_index()

        # --- PROCESAMIENTO DE REMISIONES COPIADAS Y PEGADAS ---
        df_rem = pd.read_csv(io.StringIO(datos_pegados_remisiones), sep="\t")
        df_rem.columns = df_rem.columns.astype(str).str.strip()
        
        columnas_rem_requeridas = ['Cedis', 'Ruta', 'Remision', 'Producto', 'Precio', 'Cantidad', 'SubTotal']
        for col in columnas_rem_requeridas:
            if col not in df_rem.columns:
                st.error(f"❌ Los datos pegados de las Remisiones no parecen válidos. Asegúrate de incluir los encabezados y que la columna '{col}' exista.")
                st.stop()

        # --- BASE RELACIÓN DE REMISIONES ---
        try:
            df_rel = pd.read_excel(archivo_relacion, engine='openpyxl')
        except Exception:
            try:
                df_rel = pd.read_excel(archivo_relacion, engine='xlrd')
            except Exception:
                df_rel = pd.read_html(archivo_relacion)[0]
                
        df_rel.columns = df_rel.columns.astype(str).str.strip()
        df_rel['Remision'] = df_rel['Remision'].astype(str).str.split('.').str[0].str.strip()
        df_rel['Acuse de Recibo'] = df_rel['Acuse de Recibo'].astype(str).str.split('.').str[0].str.strip()

        # --- BASE UPC ---
        df_upc = pd.read_excel(archivo_upc)
        df_upc.columns = df_upc.columns.astype(str).str.strip()
        
        df_upc['Código'] = df_upc['Código'].astype(str).str.split('.').str[0].str.strip().str.lstrip('0')
        df_upc['Código de Barras Individual'] = df_upc['Código de Barras Individual'].astype(str).str.split('.').str[0].str.strip()
        
        dict_upc = dict(zip(df_upc['Código'], df_upc['Código de Barras Individual']))
        dict_upc_inverso = dict(zip(df_upc['Código de Barras Individual'], df_upc['Código']))

        # ==========================================
        # 3. CONSTRUCCIÓN DE LA TABLA MAESTRA ORIGINAL (LALA)
        # ==========================================
        tabla_maestra = pd.DataFrame()
        
        tabla_maestra['Cedis'] = df_rem['Cedis']
        tabla_maestra['Ruta'] = pd.to_numeric(df_rem['Ruta'], errors='coerce').fillna(0).astype(int)
        
        remision_str = df_rem['Remision'].astype(str)
        tabla_maestra['Prefijo'] = remision_str.apply(lambda x: x.split('-')[0] if '-' in x else '')
        tabla_maestra['Remision'] = remision_str.apply(lambda x: x.split('-')[1] if '-' in x else x).str.strip()

        producto_str = df_rem['Producto'].astype(str)
        tabla_maestra['SKU'] = producto_str.apply(lambda x: x.split('-')[0] if '-' in x else x).str.strip().str.lstrip('0')
        tabla_maestra['Material'] = producto_str.apply(lambda x: x.split('-')[1] if '-' in x else '')

        tabla_maestra['Precio'] = df_rem['Precio']
        tabla_maestra['Cantidad'] = df_rem['Cantidad']
        tabla_maestra['SubTotal'] = df_rem['SubTotal']

        tabla_maestra = tabla_maestra.merge(df_rel[['Remision', 'Acuse de Recibo']], on='Remision', how='left')
        tabla_maestra.rename(columns={'Acuse de Recibo': 'Acuse de Recibo (SICAV)'}, inplace=True)
        tabla_maestra['Acuse de Recibo (SICAV)'] = tabla_maestra['Acuse de Recibo (SICAV)'].fillna('').astype(str).str.split('.').str[0].str.strip()

        tabla_maestra['UPC'] = tabla_maestra['SKU'].map(dict_upc).fillna('')
        tabla_maestra['Acuse de Recibo_UPC'] = tabla_maestra['Acuse de Recibo (SICAV)'] + tabla_maestra['UPC']

        dict_soriana_pzs = dict(zip(df_sor_clean['Llave_Soriana'], df_sor_clean['Cantidad Total']))
        tabla_maestra['Pzs Acuse Soriana'] = tabla_maestra['Acuse de Recibo_UPC'].map(dict_soriana_pzs).fillna(0).astype(float)
        tabla_maestra['Dif Lala - Soriana Pzs'] = tabla_maestra['Pzs Acuse Soriana'] - tabla_maestra['Cantidad'].astype(float)

        acuses_existentes_soriana = set(df_sor_clean['Folio de recibo'].unique())
        sugerencias_material = df_sor_clean.groupby('Folio de recibo')['Código de barras'].apply(list).to_dict()

        # ==========================================
        # 4. APLICACIÓN DE CONDICIONES COMPLEJAS
        # ==========================================
        def evaluar_observaciones(row):
            acuse = row['Acuse de Recibo (SICAV)']
            llave = row['Acuse de Recibo_UPC']
            dif = row['Dif Lala - Soriana Pzs']
            
            if not acuse or acuse == 'nan' or acuse == '':
                return "Sin Acuse en SICAV"
            if acuse not in acuses_existentes_soriana:
                return "Diferencia de acuse"
            if llave not in dict_soriana_pzs:
                codigos_en_ese_acuse = sugerencias_material.get(acuse, [])
                if codigos_en_ese_acuse:
                    return f"Diferencia de material (Soriana reportó: {', '.join(codigos_en_ese_acuse)})"
                return "Diferencia de material"
            if abs(dif) < 0.01:
                return ""
            else:
                return f"Diferencia en piezas ({int(dif)} pzs)"

        tabla_maestra['Observación'] = tabla_maestra.apply(evaluar_observaciones, axis=1)

        # ==========================================
        # 5. MENÚ DESPLEGABLE CON OPCIÓN "TODAS"
        # ==========================================
        lista_rutas_origen = sorted([r for r in tabla_maestra['Ruta'].unique() if r != 0])
        opciones_select = ["TODAS"] + [str(r) for r in lista_rutas_origen]
        
        ruta_seleccionada = st.selectbox("🎯 Selecciona la Ruta a Auditar:", opciones_select)
        
        if ruta_seleccionada == "TODAS":
            df_filtrado = tabla_maestra.copy()
        else:
            df_filtrado = tabla_maestra[tabla_maestra['Ruta'] == int(ruta_seleccionada)].copy()
        
        incidencias_totales = df_filtrado[df_filtrado['Observación'] != ""].shape[0]
        
        if incidencias_totales > 0:
            st.error(f"⚠️ Se identificaron {incidencias_totales} renglones con discrepancias operativas en la Ruta {ruta_seleccionada}.")
        else:
            st.success(f"✅ ¡Ruta {ruta_seleccionada} Perfecta! Todo coincide con el portal.")

        columnas_vista = [
            'Cedis', 'Ruta', 'Prefijo', 'Remision', 'SKU', 'Material', 
            'Precio', 'Cantidad', 'SubTotal', 'Acuse de Recibo (SICAV)', 
            'UPC', 'Acuse de Recibo_UPC', 'Pzs Acuse Soriana', 'Dif Lala - Soriana Pzs', 'Observación'
        ]
        
        st.dataframe(
            df_filtrado[columnas_vista], 
            use_container_width=True,
            column_config={
                "Observación": st.column_config.TextColumn("Observación", width="large")
            }
        )

        # ==========================================
        # 6. PRODUCTOS NO MAPEADOS
        # ==========================================
        st.markdown("---")
        st.markdown("### 🔍 Productos Reportados por Soriana no Mapeados / Con Incidencias Críticas")
        
        llaves_lala_correctas = set(tabla_maestra[tabla_maestra['Observación'] == '']['Acuse de Recibo_UPC'].unique())
        df_faltantes_soriana = df_sor_clean[~df_sor_clean['Llave_Soriana'].isin(llaves_lala_correctas)].copy()
        
        if ruta_seleccionada != "TODAS":
            acuses_de_ruta = set(df_filtrado['Acuse de Recibo (SICAV)'].unique())
            df_faltantes_soriana = df_faltantes_soriana[df_faltantes_soriana['Folio de recibo'].isin(acuses_de_ruta)]
        
        if not df_faltantes_soriana.empty:
            st.warning(f"Se encontraron {df_faltantes_soriana.shape[0]} registros en Soriana que presentan anomalías o descuadres.")
            
            df_faltantes_soriana['SKU (Lala)'] = df_faltantes_soriana['Código de barras'].map(dict_upc_inverso).fillna('No Encontrado')
            columnas_inferiores = ['Sucursal', 'Folio de recibo', 'SKU (Lala)', 'Código de barras', 'Descripción', 'Cantidad Total']
            
            st.dataframe(
                df_faltantes_soriana[columnas_inferiores], 
                use_container_width=True
            )
        else:
            st.success("¡Perfecto! No existen productos extra ni discrepancias en la base de Soriana.")

    except Exception as e:
        st.error(f"❌ Error al procesar la información. Detalles: {str(e)}")
else:
    st.info("💡 Por favor, completa los datos requeridos arriba (cargar los 3 archivos y pegar las remisiones) para iniciar la conciliación.")
