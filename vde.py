import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from io import BytesIO
from wordcloud import WordCloud
import re
from collections import Counter
import gspread

# ============================================
# CONFIGURACIÓN DE USUARIOS Y ACCESOS
# ============================================

USUARIOS = {
    "VDE1": {"password": "vde1", "filtro": "VDE_1"},
    "VDE2": {"password": "vde2", "filtro": "VDE_2"},
    "VDE3": {"password": "vde3", "filtro": "VDE_3"},
    "VDE4": {"password": "vde4", "filtro": "VDE_4"},
    "master": {"password": "idemefa", "filtro": None}  # None significa acceso completo
}

# ============================================
# FUNCIÓN DE AUTENTICACIÓN
# ============================================

def autenticar():
    st.sidebar.title("🔐 Acceso al Sistema")
    
    # Si ya está autenticado, mostrar información y botón de logout
    if 'autenticado' in st.session_state and st.session_state.autenticado:
        usuario_actual = st.session_state.usuario_actual
        st.sidebar.success(f"Bienvenido: {usuario_actual}")
        if st.sidebar.button("🚪 Cerrar sesión"):
            for key in ['autenticado', 'usuario_actual', 'filtro_vendedor']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
        return True
    
    # Si no está autenticado, mostrar formulario de login
    else:
        with st.sidebar.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            enviar = st.form_submit_button("Ingresar")
            
            if enviar:
                if usuario in USUARIOS and USUARIOS[usuario]["password"] == password:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = usuario
                    st.session_state.filtro_vendedor = USUARIOS[usuario]["filtro"]
                    st.rerun()
                else:
                    st.sidebar.error("Usuario o contraseña incorrectos")
        return False

# Configuración inicial de la página
st.set_page_config(
    page_title="Panel Seguimiento Ventas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .sidebar .sidebar-content {
        padding: 1rem;
    }
    .kpi-card {
        background: white;
        border-radius: 5px;
        padding: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #4CAF50;
    }
    .kpi-title {
        font-size: 0.9rem;
        color: #555;
        margin-bottom: 0.5rem;
    }
    .kpi-value {
        font-size: 1.5rem;
        font-weight: bold;
        color: #333;
    }
    .kpi-subtext {
        font-size: 0.8rem;
        color: #777;
        margin-top: 0.5rem;
    }
    .tab-content {
        padding-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Mapeo de meses inglés a español
MESES_ES = {
    'January': 'Enero',
    'February': 'Febrero',
    'March': 'Marzo',
    'April': 'Abril',
    'May': 'Mayo',
    'June': 'Junio',
    'July': 'Julio',
    'August': 'Agosto',
    'September': 'Septiembre',
    'October': 'Octubre',
    'November': 'Noviembre',
    'December': 'Diciembre'
}

MESES_ORDEN = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
               'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# ============================================
# FUNCIONES DE CARGA DE DATOS
# ============================================

@st.cache_data(ttl=3600)
def load_sales_data():
    """Carga datos de ventas desde la hoja DB_VNT"""
    # URL para DB_VNT (gid=0 es la primera hoja)
    url = "https://docs.google.com/spreadsheets/d/1WWynEjZGN8zxlIOUab1CjjMSgXhDBd4y/export?format=csv&gid=674013502"

    try:
        df = pd.read_csv(url)

        # Verificación de columnas esenciales
        required_columns = ['CODIGO', 'CLIENTE', 'Documento', 'Fecha', 'Total', 'VDE']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Faltan columnas requeridas. Columnas encontradas: {df.columns.tolist()}")
            return None

        # Limpieza y transformación
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df['MONTO'] = pd.to_numeric(df['Total'], errors='coerce')
        df['ANO'] = df['Fecha'].dt.year
        df['MES'] = df['Fecha'].dt.strftime('%B').map(MESES_ES)
        df['DIA_SEM'] = df['Fecha'].dt.day_name()
        df['SEM'] = df['Fecha'].dt.isocalendar().week

        # Orden de meses en español
        df['MES_ORDEN'] = pd.Categorical(df['MES'], categories=MESES_ORDEN, ordered=True)

        return df.dropna(subset=['Fecha', 'MONTO'])
    except Exception as e:
        st.error(f"Error al cargar datos de ventas: {str(e)}")
        return None

@st.cache_data(ttl=3600)
def load_budget_data():
    """Carga datos de presupuesto desde la hoja DB_PPTO"""
    url = "https://docs.google.com/spreadsheets/d/1WWynEjZGN8zxlIOUab1CjjMSgXhDBd4y/export?format=csv&gid=1523879888"

    try:
        df = pd.read_csv(url)

        # Verificación de columnas
        if 'MONTO' not in df.columns:
            st.error("La columna 'MONTO' no existe en DB_PPTO")
            return None

        # Limpieza de valores numéricos
        df['MONTO'] = (
            df['MONTO']
            .astype(str)
            .str.replace(r'[^\d.]', '', regex=True)  # Elimina todo excepto números y punto
            .replace('', '0')
            .astype(float)
        )

        # Convertir meses a español y manejar valores None
        if 'MES' in df.columns:
            # Primero reemplazar None o strings vacíos
            df['MES'] = df['MES'].fillna('Desconocido')
            df['MES'] = df['MES'].replace('None', 'Desconocido')

            # Mapeo de meses
            meses_map = {
                'January': 'Enero',
                'February': 'Febrero',
                'March': 'Marzo',
                'April': 'Abril',
                'May': 'Mayo',
                'June': 'Junio',
                'July': 'Julio',
                'August': 'Agosto',
                'September': 'Septiembre',
                'October': 'Octubre',
                'November': 'Noviembre',
                'December': 'Diciembre'
            }

            # Aplicar mapeo solo a valores conocidos
            df['MES'] = df['MES'].apply(lambda x: meses_map.get(x, x))

            # Reemplazar 'Desconocido' con el mes correspondiente a la fecha si existe columna Fecha
            if 'Fecha' in df.columns:
                df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
                df.loc[df['MES'] == 'Desconocido', 'MES'] = df['Fecha'].dt.strftime('%B').map(meses_map)

        return df
    except Exception as e:
        st.error(f"Error al cargar datos de presupuesto: {str(e)}")
        return None

@st.cache_data(ttl=3600)
def load_clients_data():
    """Carga datos de clientes desde la hoja DB_CLI"""
    # URL para DB_CLI (necesitas el gid correcto)
    url = "https://docs.google.com/spreadsheets/d/1WWynEjZGN8zxlIOUab1CjjMSgXhDBd4y/export?format=csv&gid=81018902"

    try:
        df = pd.read_csv(url)

        # Verificación de columnas esenciales
        if 'CODIGO' not in df.columns or 'CLIENTE' not in df.columns:
            st.error("Faltan columnas requeridas (CODIGO o CLIENTE)")
            return None

        return df
    except Exception as e:
        st.error(f"Error al cargar datos de clientes: {str(e)}")
        return None

@st.cache_data(ttl=3600)
def load_calls_data():
    # URL para el archivo de análisis de audio (Sheet1)
    url = "https://docs.google.com/spreadsheets/d/1Zlpdoq-lV8dpF1jITGlAHXRtmrJHhMm0/export?format=csv"

    try:
        df = pd.read_csv(url)

        # Limpieza y transformación de datos
        df.columns = df.columns.str.strip()
        df["Duración (min)"] = df["Duración (seg)"] / 60

        try:
            df["Fecha"] = pd.to_datetime(df["Archivo"].str.extract(r"(\d{4}-\d{2}-\d{2})")[0], errors="coerce")
            df["MES"] = df["Fecha"].dt.strftime('%B').map(MESES_ES)
        except:
            df["Fecha"] = pd.NaT
            df["MES"] = ""

        mapa_fluidez = {"Excelente": 10, "Buena": 8, "Regular": 6, "Deficiente": 4, "Malo": 2}
        df["Puntaje Fluidez"] = df["Evaluación Fluidez"].map(mapa_fluidez).fillna(0)

        df["Puntaje Calidad"] = (df["% Apego al guion"] * 0.5) + \
                                (df["% Sentimiento"] * 0.3) + \
                                (df["Puntaje Fluidez"] * 10 * 0.2)

        def clasificar(p):
            if p >= 85: return "🏆 Ejemplar"
            elif p >= 70: return "✅ Satisfactorio"
            elif p >= 50: return "⚠️ Necesita Mejora"
            else: return "🔴 Crítico"

        df["Clasificación"] = df["Puntaje Calidad"].apply(clasificar)

        return df
    except Exception as e:
        st.error(f"Error al cargar datos de llamadas: {str(e)}")
        return None

# ============================================
# CLASE PARA CÁLCULOS DE KPIs
# ============================================

class KPICalculator:
    def __init__(self, sales_df, budget_df):
        self.sales_df = sales_df
        self.budget_df = budget_df

    def calcular_cumplimiento_metas(self, df_filtrado, mes, year):
        try:
            # Filtrar presupuesto por mes y año
            presupuesto_mes = self.budget_df[
                (self.budget_df['MES'] == mes) &
                (self.budget_df['ANO'] == year)
            ].groupby('VDE').agg({'MONTO': 'sum', 'CANTIDAD': 'sum'}).reset_index()

            # Si no hay presupuesto para este mes/año
            if presupuesto_mes.empty or presupuesto_mes['MONTO'].sum() == 0:
                # Verificar si hay ventas para ese mes
                ventas_mes = df_filtrado[
                    (df_filtrado['MES'] == mes) &
                    (df_filtrado['ANO'] == year)
                ]

                if not ventas_mes.empty:
                    st.warning(f"No hay datos de presupuesto para {mes} {year}, pero sí existen ventas registradas")
                else:
                    st.warning(f"No hay datos de presupuesto ni ventas para {mes} {year}")
                return None

            # Filtrar ventas por mes y año
            ventas_mes = df_filtrado[
                (df_filtrado['MES'] == mes) &
                (df_filtrado['ANO'] == year)
            ].groupby('VDE').agg({
                'MONTO': 'sum',
                'Cantidad': 'sum',
                'Documento': 'nunique',
                'CODIGO': 'nunique'
            }).reset_index()

            # Combinar datos
            cumplimiento = pd.merge(
                ventas_mes,
                presupuesto_mes,
                on='VDE',
                how='outer',
                suffixes=('_real', '_meta')
            ).fillna(0)

            # Calcular porcentajes de cumplimiento
            cumplimiento['% Cumplimiento Ventas'] = (cumplimiento['MONTO_real'] / cumplimiento['MONTO_meta']) * 100
            cumplimiento['% Cumplimiento Cajas'] = (cumplimiento['Cantidad'] / cumplimiento['CANTIDAD']) * 100
            cumplimiento['Ticket Promedio'] = cumplimiento['MONTO_real'] / cumplimiento['Documento']
            cumplimiento['Factura Promedio'] = cumplimiento['MONTO_real'] / cumplimiento['CODIGO']

            # Agregar fila de totales
            total_row = cumplimiento.sum(numeric_only=True)
            total_row['VDE'] = 'TOTAL'
            total_row['% Cumplimiento Ventas'] = (total_row['MONTO_real'] / total_row['MONTO_meta']) * 100
            total_row['% Cumplimiento Cajas'] = (total_row['Cantidad'] / total_row['CANTIDAD']) * 100
            total_row['Ticket Promedio'] = total_row['MONTO_real'] / total_row['Documento']
            total_row['Factura Promedio'] = total_row['MONTO_real'] / total_row['CODIGO']

            cumplimiento = pd.concat([cumplimiento, pd.DataFrame([total_row])], ignore_index=True)

            return {
                'dataframe': cumplimiento,
                'mes': mes,
                'year': year
            }
        except Exception as e:
            st.error(f"Error al calcular cumplimiento: {str(e)}")
            return None

    def calcular_kpis_basicos(self, df):
        try:
            kpis = {
                'clientes_unicos': df['CODIGO'].nunique(),
                'transacciones': df['Documento'].nunique(),
                'frecuencia_compra': df.groupby('CODIGO')['Documento'].nunique().mean(),
                'ticket_promedio': df['MONTO'].sum() / df['Documento'].nunique(),
                'factura_promedio': df['MONTO'].sum() / df['CODIGO'].nunique()
            }
            return kpis
        except Exception as e:
            st.error(f"Error al calcular KPIs básicos: {str(e)}")
            return None

    def calcular_proyeccion_semanal(self, df, semana):
        try:
            df_semana = df[df['SEM'] == semana]

            if df_semana.empty:
                return None

            hoy = datetime.now().date()
            primer_dia_semana = hoy - timedelta(days=hoy.weekday())
            ultimo_dia_semana = primer_dia_semana + timedelta(days=6)

            # Calcular días laborables (considerando sábado medio día)
            dias_transcurridos = min((hoy - primer_dia_semana).days + 1, 5)  # Máximo 5 días laborables
            if hoy.weekday() == 5:  # Si es sábado
                dias_transcurridos = 5.5
            elif hoy.weekday() == 6:  # Si es domingo
                dias_transcurridos = 5

            dias_laborables = 5.5  # Semana laboral de lunes a sábado medio día

            ventas_semana = df_semana['MONTO'].sum()
            venta_diaria_promedio = ventas_semana / dias_transcurridos if dias_transcurridos > 0 else 0
            proyeccion = venta_diaria_promedio * dias_laborables

            return {
                'semana': semana,
                'ventas_semana': ventas_semana,
                'dias_transcurridos': dias_transcurridos,
                'dias_laborables': dias_laborables,
                'venta_diaria_promedio': venta_diaria_promedio,
                'proyeccion': proyeccion
            }
        except Exception as e:
            st.error(f"Error al calcular proyección semanal: {str(e)}")
            return None

    def calcular_proyeccion_mensual(self, df, mes):
        try:
            df_mes = df[df['MES'] == mes]

            if df_mes.empty:
                return None

            hoy = datetime.now().date()
            primer_dia_mes = hoy.replace(day=1)
            ultimo_dia_mes = (primer_dia_mes + timedelta(days=32)).replace(day=1) - timedelta(days=1)

            # Calcular días laborables (considerando sábados medio día)
            def calcular_dias_laborables(start, end):
                dias = 0
                current = start
                while current <= end:
                    if current.weekday() < 5:  # Lunes a Viernes
                        dias += 1
                    elif current.weekday() == 5:  # Sábado
                        dias += 0.5
                    current += timedelta(days=1)
                return dias

            dias_transcurridos = calcular_dias_laborables(primer_dia_mes, hoy)
            dias_totales = calcular_dias_laborables(primer_dia_mes, ultimo_dia_mes)

            ventas_mes = df_mes['MONTO'].sum()
            venta_diaria_promedio = ventas_mes / dias_transcurridos if dias_transcurridos > 0 else 0
            proyeccion = venta_diaria_promedio * dias_totales

            ticket_promedio = df_mes['MONTO'].sum() / df_mes['Documento'].nunique() if df_mes['Documento'].nunique() > 0 else 0
            factura_promedio = df_mes['MONTO'].sum() / df_mes['CODIGO'].nunique() if df_mes['CODIGO'].nunique() > 0 else 0

            return {
                'mes': mes,
                'ventas_mes': ventas_mes,
                'dias_transcurridos': dias_transcurridos,
                'dias_totales': dias_totales,
                'venta_diaria_promedio': venta_diaria_promedio,
                'proyeccion': proyeccion,
                'ticket_promedio': ticket_promedio,
                'factura_promedio': factura_promedio
            }
        except Exception as e:
            st.error(f"Error al calcular proyección mensual: {str(e)}")
            return None

# ============================================
# FUNCIONES AUXILIARES
# ============================================

def format_monto(value):
    return f"${value:,.2f}"

def format_cantidad(value):
    return f"{value:,.0f}"

def crear_card(title, value, value_type='monto'):
    if value_type == 'monto':
        value_str = format_monto(value)
    else:
        value_str = format_cantidad(value)

    return f"""
    <div class="kpi-card">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value_str}</div>
    </div>
    """

# ============================================
# INTERFAZ PRINCIPAL
# ============================================

def main():
    # Verificar autenticación primero
    if not autenticar():
        st.warning("Por favor ingrese sus credenciales en la barra lateral")
        return
    
    # Obtener el filtro del usuario actual
    usuario_actual = st.session_state.usuario_actual
    filtro_vendedor = st.session_state.filtro_vendedor
    
    st.title(f"📊 Panel de Control Comercial - {usuario_actual}")

    # Cargar datos
    sales_df = load_sales_data()
    budget_df = load_budget_data()
    calls_df = load_calls_data()

    if sales_df is None or budget_df is None or calls_df is None:
        st.error("No se pudieron cargar los datos necesarios. Por favor intente más tarde.")
        return

    # Aplicar filtro de vendedor si corresponde (excepto para master)
    if filtro_vendedor:
        sales_df = sales_df[sales_df['VDE'] == filtro_vendedor]
        budget_df = budget_df[budget_df['VDE'] == filtro_vendedor]
        calls_df = calls_df[calls_df['Vendedor'] == filtro_vendedor]
        st.info(f"Mostrando información solo para el vendedor: {filtro_vendedor}")

    # Crear instancia del calculador de KPIs
    kpi_calculator = KPICalculator(sales_df, budget_df)

    # Barra de navegación superior (módulos)
    st.sidebar.title("Módulos")
    module = st.sidebar.radio("Seleccione módulo:", ["Consulta", "Llamadas", "Cumplimiento", "Proyecciones"])

    # Filtros globales en el sidebar
    st.sidebar.header("🔎 Filtros Globales")

    # Año (común a varios módulos)
    available_years = sorted(sales_df['ANO'].unique(), reverse=True)
    year_filter = st.sidebar.selectbox("Año", available_years)

    # Módulo de Consulta
    if module == "Consulta":
        st.header("📊 Consulta de Ventas por Producto")

        # Barra superior con controles de actualización
        col1, col2, col3 = st.columns([6, 1, 1])
        with col1:
            st.write("")  # Espacio para alinear
        with col2:
            if st.button("🔄 Recargar Datos", help="Actualizar datos desde Google Drive"):
                st.cache_data.clear()  # Limpiar caché para forzar recarga
        with col3:
            last_update = st.empty()  # Espacio reservado para mostrar última actualización

        # Sidebar para filtros específicos de consulta
        with st.sidebar:
            st.header("🔎 Filtros de Consulta")

            search_option = st.radio("Buscar por:", ["Código", "Descripción", "Cliente"])

            if search_option == "Código":
                codigos = sorted(sales_df['COD_PROD'].unique())
                cod_input = st.selectbox("Seleccione código de producto", codigos)
            elif search_option == "Descripción":
                descripciones = sorted(sales_df['Descripcion'].unique())
                desc_selected = st.selectbox("Seleccione descripción", descripciones)
                cod_input = sales_df[sales_df['Descripcion'] == desc_selected]['COD_PROD'].iloc[0]
            else:
                clientes = sorted(sales_df['CLIENTE'].unique())
                cliente_sel = st.selectbox("Seleccione cliente", clientes)
                cod_input = None

            min_date = sales_df['Fecha'].min().date() if not sales_df['Fecha'].isnull().all() else pd.to_datetime('today').date()
            max_date = sales_df['Fecha'].max().date() if not sales_df['Fecha'].isnull().all() else pd.to_datetime('today').date()

            col1, col2 = st.columns(2)
            with col1:
                fecha_inicio = st.date_input("Desde", min_date)
            with col2:
                fecha_fin = st.date_input("Hasta", max_date)

            vendedores = sorted(sales_df['VDE'].unique())
            vendedores_sel = st.multiselect("Vendedor(es)", vendedores)

            group_by = st.selectbox("Agrupar por", ["Ninguno", "Vendedor", "Cliente", "Mes", "Año"])

        # Aplicar filtros
        if search_option == "Cliente":
            mask = (
                (sales_df['CLIENTE'] == cliente_sel) &
                (sales_df['Fecha'].dt.date >= fecha_inicio) &
                (sales_df['Fecha'].dt.date <= fecha_fin)
            )
            titulo = f"Ventas para el cliente: {cliente_sel}"
        else:
            mask = (
                (sales_df['COD_PROD'] == cod_input) &
                (sales_df['Fecha'].dt.date >= fecha_inicio) &
                (sales_df['Fecha'].dt.date <= fecha_fin)
            )
            producto = sales_df[sales_df['COD_PROD'] == cod_input]['Descripcion'].iloc[0]
            titulo = f"Ventas para: {cod_input} - {producto}"

        if vendedores_sel:
            mask &= sales_df['VDE'].isin(vendedores_sel)

        resultado = sales_df[mask].copy().sort_values('Fecha')

        if not resultado.empty:
            resultado['Fecha_mostrar'] = resultado['Fecha'].dt.strftime('%d/%m/%Y')

            st.subheader(titulo)

            # Agrupación de datos
            if group_by != "Ninguno":
                if group_by == "Mes":
                    resultado['Grupo'] = resultado['Fecha'].dt.to_period('M').astype(str)
                elif group_by == "Año":
                    resultado['Grupo'] = resultado['Fecha'].dt.year.astype(str)
                elif group_by == "Vendedor":
                    resultado['Grupo'] = resultado['VDE']
                else:  # Cliente
                    resultado['Grupo'] = resultado['CLIENTE']

                grouped = resultado.groupby('Grupo').agg({
                    'Cantidad': 'sum',
                    'MONTO': 'sum',
                    'Documento': 'nunique'
                }).reset_index()
                grouped.rename(columns={'Documento': 'Transacciones'}, inplace=True)

                st.subheader(f"📊 Ventas agrupadas por {group_by.lower()}")
                st.dataframe(grouped)

                # Gráfico de barras
                if not grouped.empty:
                    fig = px.bar(
                        grouped,
                        x='Grupo',
                        y='MONTO',
                        text='MONTO',
                        title=f"Ventas por {group_by.lower()}",
                        labels={'MONTO': 'Monto Total', 'Grupo': group_by},
                        hover_data=['Cantidad', 'Transacciones']
                    )
                    fig.update_traces(
                        texttemplate='%{text:,.2f}',
                        textposition='outside',
                        marker_color='#4CAF50'
                    )
                    fig.update_layout(
                        xaxis_title=group_by,
                        yaxis_title="Monto Total",
                        height=500
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # Detalle de transacciones
            st.subheader("📋 Detalle de transacciones")
            columnas_mostrar = [
                'CLIENTE', 'VDE', 'Fecha_mostrar', 'Documento',
                'Descripcion', 'Cantidad', 'MONTO'
            ]
            st.dataframe(
                resultado[columnas_mostrar].rename(columns={
                    'Fecha_mostrar': 'Fecha',
                    'MONTO': 'Monto'
                })
            )

            # Métricas
            total_cant = resultado['Cantidad'].sum()
            total_monto = resultado['MONTO'].sum()
            transacciones = resultado['Documento'].nunique()
            avg_price = total_monto / total_cant if total_cant > 0 else 0
            ticket_promedio = total_monto / transacciones if transacciones > 0 else 0

            st.subheader("📊 Totales")
            col1, col2, col3, col4, col5 = st.columns(5)

            col1.metric("Total Unidades", f"{total_cant:,.0f}")
            col2.metric("Total Ventas", f"${total_monto:,.2f}")
            col3.metric("Precio Promedio", f"${avg_price:,.2f}")
            col4.metric("Ticket Promedio", f"${ticket_promedio:,.2f}")
            col5.metric("Transacciones", f"{transacciones:,.0f}")

            # Gráfico de línea
            if len(resultado) > 1:
                fig = px.line(
                    resultado,
                    x='Fecha',
                    y='MONTO',
                    title='Evolución de Ventas por Fecha',
                    markers=True,
                    labels={'MONTO': 'Monto', 'Fecha': 'Fecha'},
                    hover_data=['CLIENTE', 'VDE', 'Cantidad']
                )
                fig.update_layout(
                    xaxis_title="Fecha",
                    yaxis_title="Monto",
                    height=500
                )
                fig.update_traces(line_color='#FF4B4B', marker_color='#FF4B4B')
                st.plotly_chart(fig, use_container_width=True)

            # Exportación
            st.subheader("💾 Exportar Resultados")
            export_format = st.radio("Formato de exportación:", ["Excel", "CSV"])

            try:
                if export_format == "Excel":
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        resultado.drop(columns=['Fecha_mostrar']).to_excel(
                            writer, index=False, sheet_name='Detalle')
                        if group_by != "Ninguno":
                            grouped.to_excel(writer, index=False, sheet_name='Agrupado')
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=output.getvalue(),
                        file_name="reporte_ventas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.download_button(
                        label="⬇️ Descargar CSV",
                        data=resultado.drop(columns=['Fecha_mostrar'])
                            .to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8'),
                        file_name="reporte_ventas.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Error al exportar: {str(e)}")
                st.info("ℹ️ Si el error persiste, intente exportar como CSV o instale xlsxwriter manualmente")

        else:
            st.warning("⚠️ No se encontraron resultados con los filtros aplicados")

        # Mostrar última actualización
        last_update.caption(f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    # Módulo de Llamadas
    elif module == "Llamadas":
        st.header("📞 Análisis de Llamadas")

        # Filtros específicos de llamadas
        with st.sidebar:
            st.header("Filtros de Llamadas")
            vendedores = st.multiselect(
                "Seleccionar vendedor",
                options=calls_df["Vendedor"].unique(),
                default=calls_df["Vendedor"].unique()
            )

            fecha_min = st.date_input(
                "Fecha mínima",
                value=calls_df["Fecha"].min() if calls_df["Fecha"].notna().any() else None
            )
            fecha_max = st.date_input(
                "Fecha máxima",
                value=calls_df["Fecha"].max() if calls_df["Fecha"].notna().any() else None
            )

        # Aplicar filtros
        df_filtrado = calls_df.copy()
        if vendedores:
            df_filtrado = df_filtrado[df_filtrado["Vendedor"].isin(vendedores)]
        if fecha_min and fecha_max:
            df_filtrado = df_filtrado[
                (df_filtrado["Fecha"] >= pd.to_datetime(fecha_min)) &
                (df_filtrado["Fecha"] <= pd.to_datetime(fecha_max))
            ]

        # Pestañas para diferentes vistas
        tab1, tab2, tab3 = st.tabs(["📊 Resumen general", "👤 Vista por vendedor", "🗣 Análisis de lenguaje"])

        with tab1:
            st.subheader("Resumen general del equipo")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Promedio Apego (%)", f"{df_filtrado['% Apego al guion'].mean():.1f}%")
            col2.metric("Promedio Sentimiento (%)", f"{df_filtrado['% Sentimiento'].mean():.1f}%")
            col3.metric("Llamadas analizadas", len(df_filtrado))
            col4.metric("Duración promedio (min)", f"{df_filtrado['Duración (min)'].mean():.1f}")

            # Comparativa por vendedor
            fig_bar = px.bar(
                df_filtrado.groupby("Vendedor")["% Apego al guion"].mean().reset_index(),
                x="Vendedor",
                y="% Apego al guion",
                title="Apego al guion por vendedor",
                color="% Apego al guion",
                color_continuous_scale="RdYlGn"
            )
            st.plotly_chart(fig_bar, use_container_width=True)

            # Tabla de calidad promedio por vendedor
            st.markdown("### 🏅 Calidad promedio por vendedor")
            calidad_vendedores = df_filtrado.groupby("Vendedor")[["Puntaje Calidad"]].mean().reset_index()
            calidad_vendedores["Clasificación"] = calidad_vendedores["Puntaje Calidad"].apply(
                lambda x: df_filtrado["Clasificación"].iloc[0] if len(df_filtrado) else ""
            )
            st.dataframe(calidad_vendedores)

        with tab2:
            vendedor_sel = st.selectbox(
                "Seleccionar vendedor",
                options=df_filtrado["Vendedor"].unique()
            )
            df_vend = df_filtrado[df_filtrado["Vendedor"] == vendedor_sel]

            st.subheader(f"Desempeño de {vendedor_sel}")

            # Gauge Apego
            if not df_vend.empty:
                apego = df_vend["% Apego al guion"].mean()
                fig_gauge = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=apego,
                    title={'text': "Apego al guion (%)"},
                    gauge={'axis': {'range': [0, 100]},
                           'bar': {'color': "green" if apego >= 80 else "orange" if apego >= 50 else "red"}}
                ))
                st.plotly_chart(fig_gauge, use_container_width=True)

            # Tabla con clasificación por llamada
            st.markdown("### 📋 Calidad de llamadas")
            st.dataframe(df_vend[[
                "Archivo", "Duración (min)", "% Apego al guion",
                "% Sentimiento", "Puntaje Calidad", "Clasificación"
            ]])

        with tab3:
            st.subheader("Análisis de lenguaje")

            # Nube de palabras
            texto_completo = " ".join(df_filtrado["Transcripción completa"].dropna().astype(str))
            if texto_completo:
                wc = WordCloud(width=800, height=400, background_color="white").generate(texto_completo)
                st.image(wc.to_array(), caption="Nube de palabras", use_container_width=True)

            # Palabras de relleno
            if "Transcripción completa" in df_filtrado.columns:
                palabras_relleno = re.findall(r"\b(eh|este|o sea|mmm|ah)\b", texto_completo.lower())
                conteo_relleno = Counter(palabras_relleno)
                df_relleno = pd.DataFrame(
                    conteo_relleno.items(),
                    columns=["Palabra", "Frecuencia"]
                ).sort_values(by="Frecuencia", ascending=False)

                st.markdown("### 📊 Distribución de palabras de relleno")
                fig_fill = px.bar(
                    df_relleno,
                    x="Palabra",
                    y="Frecuencia",
                    title="Frecuencia de palabras de relleno"
                )
                st.plotly_chart(fig_fill, use_container_width=True)
                st.dataframe(df_relleno)

            # Comparativa de energía y tono
            if "Energía de voz" in df_filtrado.columns and "Tono promedio" in df_filtrado.columns:
                fig_scatter = px.scatter(
                    df_filtrado,
                    x="Tono promedio",
                    y="Energía de voz",
                    color="% Apego al guion",
                    size="Tasa de habla",
                    title="Tono vs Energía"
                )
                st.plotly_chart(fig_scatter, use_container_width=True)

    # Módulo de Cumplimiento
    elif module == "Cumplimiento":
        st.header("🎯 Cumplimiento de Metas")

        # Filtros específicos de cumplimiento
        with st.sidebar:
            st.header("Filtros de Cumplimiento")

            # Selección de meses - ahora ordenados cronológicamente
            meses_disponibles = sorted(sales_df['MES'].unique(), key=lambda x: MESES_ORDEN.index(x))
            meses_seleccionados = st.multiselect(
                "Seleccionar mes(es)",
                options=meses_disponibles,
                default=[meses_disponibles[0]] if meses_disponibles else []
            )

            # Selección de vendedores
            vendedores_disponibles = sorted(sales_df['VDE'].unique())
            vendedores_seleccionados = st.multiselect(
                "Seleccionar vendedor(es)",
                options=vendedores_disponibles,
                default=vendedores_disponibles
            )

        if not meses_seleccionados:
            st.warning("Por favor seleccione al menos un mes para analizar")
            return

        # Filtrar datos por año y vendedores seleccionados
        current_df = sales_df[
            (sales_df['ANO'] == year_filter) &
            (sales_df['VDE'].isin(vendedores_seleccionados))
        ]

        if current_df.empty:
            st.warning("No hay datos disponibles con los filtros seleccionados")
            return

        # Mostrar cumplimiento para el primer mes seleccionado
        mes_analisis = meses_seleccionados[0]
        cumplimiento = kpi_calculator.calcular_cumplimiento_metas(current_df, mes_analisis, year_filter)

        if cumplimiento:
            df_renombrado = cumplimiento['dataframe'].rename(columns={
                'MONTO_real': 'MONTO REAL',
                'MONTO_meta': 'MONTO META',
                'Cantidad': 'CANTIDAD REAL',
                'CANTIDAD': 'CANTIDAD META',
                'VDE': 'VENDEDOR',
                'Documento': 'FACTURAS',
                'CODIGO': 'CLIENTES',
                '% Cumplimiento Ventas': '% CUMPL. VENTAS',
                '% Cumplimiento Cajas': '% CUMPL. CANTIDAD'
            })

            st.dataframe(
                df_renombrado.style.format({
                    'MONTO REAL': "${:,.2f}",
                    'MONTO META': "${:,.2f}",
                    '% CUMPL. VENTAS': "{:.1f}%",
                    'CANTIDAD REAL': "{:,.0f}",
                    'CANTIDAD META': "{:,.0f}",
                    '% CUMPL. CANTIDAD': "{:.1f}%",
                    'CLIENTES': "{:,}",
                    'FACTURAS': "{:,}",
                    'Ticket Promedio': "${:,.2f}",
                    'Factura Promedio': "${:,.2f}"
                }).map(
                    lambda x: 'color: green' if isinstance(x, (int, float)) and x >= 100 else
                            'color: orange' if isinstance(x, (int, float)) and x >= 70 else
                            'color: red' if isinstance(x, (int, float)) and x < 70 else '',
                    subset=['% CUMPL. VENTAS', '% CUMPL. CANTIDAD']
                ),
                use_container_width=True
            )

            # SECCIÓN DE ESFUERZO DIARIO REQUERIDO
            st.subheader("📊 Esfuerzo Diario Requerido para Alcanzar Metas")

            # Usamos el dataframe renombrado para mayor claridad
            resumen = df_renombrado

            # Verificar columnas disponibles
            required_cols = {
                'monto_real': 'MONTO REAL',
                'monto_meta': 'MONTO META',
                'cantidad_real': 'CANTIDAD REAL',
                'cantidad_meta': 'CANTIDAD META',
                'documentos': 'FACTURAS'
            }

            # Verificar si todas las columnas necesarias están presentes
            missing_cols = [col for col in required_cols.values() if col not in resumen.columns]

            if missing_cols:
                st.warning(f"No se pueden calcular los requerimientos diarios. Faltan columnas: {', '.join(missing_cols)}")
                st.write("Columnas disponibles:", resumen.columns.tolist())
            else:
                total_row = resumen.iloc[-1]
                hoy = datetime.now().date()
                primer_dia_mes = hoy.replace(day=1)
                ultimo_dia_mes = (primer_dia_mes + timedelta(days=32)).replace(day=1) - timedelta(days=1)

                def calcular_dias_laborables(start, end):
                    return sum(
                        0.5 if (start + timedelta(days=i)).weekday() == 5 else
                        1 if (start + timedelta(days=i)).weekday() != 6 else 0
                        for i in range((end - start).days + 1)
                    )

                dias_transcurridos = calcular_dias_laborables(primer_dia_mes, hoy)
                dias_totales = calcular_dias_laborables(primer_dia_mes, ultimo_dia_mes)
                dias_faltantes = max(0, dias_totales - dias_transcurridos)
                clientes_meta_total = 25 * dias_totales

                if dias_faltantes > 0:
                    # METAS
                    meta_monto = total_row[required_cols['monto_meta']]
                    meta_cantidad = total_row[required_cols['cantidad_meta']]
                    meta_clientes_dia = 25
                    meta_ticket = meta_monto / meta_cantidad if meta_cantidad > 0 else 0
                    meta_factura = meta_monto / clientes_meta_total if clientes_meta_total > 0 else 0

                    # REALES
                    real_monto = total_row[required_cols['monto_real']]
                    real_cantidad = total_row[required_cols['cantidad_real']]
                    real_clientes = total_row[required_cols['documentos']]
                    real_clientes_dia = real_clientes / dias_transcurridos if dias_transcurridos > 0 else 0
                    real_ticket = real_monto / real_cantidad if real_cantidad > 0 else 0
                    real_factura = real_monto / real_clientes if real_clientes > 0 else 0

                    # REQUERIMIENTO
                    ventas_faltantes = max(0, meta_monto - real_monto)
                    ventas_diarias_requeridas = ventas_faltantes / dias_faltantes

                    cantidad_faltantes = max(0, meta_cantidad - real_cantidad)
                    cantidad_diarias_requeridas = cantidad_faltantes / dias_faltantes

                    clientes_faltantes = max(0, clientes_meta_total - real_clientes)
                    clientes_diarios_requeridos = clientes_faltantes / dias_faltantes if dias_faltantes > 0 else 0

                    ticket_requerido = ventas_diarias_requeridas / cantidad_diarias_requeridas if cantidad_diarias_requeridas > 0 else 0
                    factura_requerida = ventas_diarias_requeridas / clientes_diarios_requeridos if clientes_diarios_requeridos > 0 else 0

                    # TARJETAS
                    cols = st.columns(3)

                    with cols[0]:  # META
                        st.markdown(f"""
                        <div class="kpi-card" style="border-left-color: #4CAF50">
                            <div class="kpi-title">🎯 Meta Mensual</div>
                            <div class="kpi-value">{format_monto(meta_monto)}</div>
                            <div class="kpi-subtext">
                            <div><span class="kpi-value">👥 Clientes/día: {meta_clientes_dia:.1f}</span></div>
                            <div><span class="kpi-value">📦 Cantidad/día: {format_cantidad(meta_cantidad/dias_totales)}</span></div>
                            <div><span class="kpi-value">🎫 Ticket: {format_monto(meta_ticket)}</span></div>
                            <div><span class="kpi-value">🧾 Factura: {format_monto(meta_factura)}</span></div>
                        </div>
                        """, unsafe_allow_html=True)

                    with cols[1]:  # REAL
                        st.markdown(f"""
                        <div class="kpi-card">
                            <div class="kpi-title">📌 Real Acumulado</div>
                            <div class="kpi-value">{format_monto(real_monto)}</div>
                            <div class="kpi-subtext">
                                <div><span class="kpi-value">👥 Clientes/día: <span class="kpi-value">{real_clientes_dia:.1f}</span><br>
                                <div><span class="kpi-value">📦 Cantidad: <span class="kpi-value">{format_cantidad(real_cantidad)}</span><br>
                                <div><span class="kpi-value">🎫 Ticket: <span class="kpi-value">{format_monto(real_ticket)}</span><br>
                                <div><span class="kpi-value">🧾 Factura: <span class="kpi-value">{format_monto(real_factura)}</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    with cols[2]:  # ESFUERZO
                        st.markdown(f"""
                        <div class="kpi-card" style="border-left-color: #FFA500">
                            <div class="kpi-title">⚡ Esfuerzo Diario</div>
                            <div class="kpi-value">{format_monto(ventas_diarias_requeridas)}</div>
                            <div class="kpi-subtext">
                                <div><span class="kpi-value">👥 Clientes/día: <span class="kpi-value">{clientes_diarios_requeridos:.1f}</span><br>
                                <div><span class="kpi-value">📦 Cantidad/día: <span class="kpi-value">{format_cantidad(cantidad_diarias_requeridas)}</span><br>
                                <div><span class="kpi-value">🎫 Ticket: <span class="kpi-value">{format_monto(ticket_requerido)}</span><br>
                                <div><span class="kpi-value">🧾 Factura: <span class="kpi-value">{format_monto(factura_requerida)}</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                else:
                    st.warning("El mes ha concluido. No quedan días laborables para planificar.")

            # SECCIÓN ANALÍTICA
            st.subheader("🔍 Analítica Comercial")

            # KPIs básicos usando el método de la clase KPICalculator
            kpis_basicos = kpi_calculator.calcular_kpis_basicos(current_df)

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.markdown(crear_card(
                    "Clientes Totales",
                    kpis_basicos['clientes_unicos'],
                    'cantidad'
                ), unsafe_allow_html=True)
            with col2:
                st.markdown(crear_card(
                    "Transacciones",
                    kpis_basicos['transacciones'],
                    'cantidad'
                ), unsafe_allow_html=True)
            with col3:
                st.markdown(crear_card(
                    "Frecuencia Compra",
                    kpis_basicos['frecuencia_compra'],
                    'cantidad'
                ), unsafe_allow_html=True)
            with col4:
                st.markdown(crear_card(
                    "Ticket Promedio",
                    kpis_basicos['ticket_promedio'],
                    'monto'
                ), unsafe_allow_html=True)
            with col5:
                st.markdown(crear_card(
                    "Factura Promedio",
                    kpis_basicos['factura_promedio'],
                    'monto'
                ), unsafe_allow_html=True)

            # Tabla de Ventas por Vendedor y Mes - Versión corregida con orden de meses
            st.markdown("### 📊 Ventas por Vendedor y Mes")

            # Crear un DataFrame pivote con los meses ordenados
            ventas_por_mes = current_df.pivot_table(
                index='VDE',
                columns='MES',
                values='MONTO',
                aggfunc='sum',
                fill_value=0
            )

            # Reordenar las columnas según MESES_ORDEN, manteniendo solo los meses existentes
            meses_existentes = [mes for mes in MESES_ORDEN if mes in ventas_por_mes.columns]
            ventas_por_mes = ventas_por_mes[meses_existentes]

            # Formatear la salida
            styled_ventas = ventas_por_mes.style.format("${:,.2f}").background_gradient(cmap='Blues')
            st.dataframe(styled_ventas, use_container_width=True)

            # Cumplimiento por Categoría (Versión corregida)
            st.markdown("### 🎯 Cumplimiento por Categoría")
            try:
                # Primero unimos con los datos de presupuesto si es necesario
                if 'CATEGORIA' in current_df.columns and 'CATEGORIA' in budget_df.columns:
                    # Agrupamos ventas por categoría
                    ventas_categoria = current_df.groupby(['VDE', 'CATEGORIA'])['MONTO'].sum().reset_index()
                    
                    # Agrupamos presupuesto por categoría (asumiendo que existe en budget_df)
                    presupuesto_categoria = budget_df.groupby(['VDE', 'CATEGORIA'])['MONTO'].sum().reset_index()
                    
                    # Unimos los datos
                    cumplimiento_categoria = pd.merge(
                        ventas_categoria,
                        presupuesto_categoria,
                        on=['VDE', 'CATEGORIA'],
                        how='left',
                        suffixes=('_Real', '_Presupuesto')
                    )
                    
                    # Calculamos el cumplimiento
                    cumplimiento_categoria['% Cumplimiento'] = (cumplimiento_categoria['MONTO_Real'] / 
                                                            cumplimiento_categoria['MONTO_Presupuesto']) * 100
                    
                    # Formatear y mostrar
                    st.dataframe(
                        cumplimiento_categoria.style.format({
                            'MONTO_Real': '${:,.2f}',
                            'MONTO_Presupuesto': '${:,.2f}',
                            '% Cumplimiento': '{:.1f}%'
                        }).bar(subset=['% Cumplimiento'], align='mid', color=['#d65f5f', '#5fba7d']),
                        use_container_width=True
                    )
                else:
                    st.warning("No se encontró la columna 'CATEGORIA' en los datos")
            except Exception as e:
                st.error(f"No se pudo calcular el cumplimiento por categoría: {str(e)}")
                st.write("Columnas disponibles en ventas:", current_df.columns.tolist())
                st.write("Columnas disponibles en presupuesto:", budget_df.columns.tolist())

            # Cumplimiento por Subcategoría (Versión corregida)
            st.markdown("### 🎯 Cumplimiento por Subcategoría")
            try:
                if all(col in current_df.columns for col in ['VDE', 'CATEGORIA', 'SUBCATEGORIA']):
                    # Agrupamos ventas por subcategoría
                    ventas_subcat = current_df.groupby(['VDE', 'CATEGORIA', 'SUBCATEGORIA'])['MONTO'].sum().reset_index()
                    
                    if all(col in budget_df.columns for col in ['VDE', 'CATEGORIA', 'SUBCATEGORIA']):
                        # Agrupamos presupuesto por subcategoría
                        presupuesto_subcat = budget_df.groupby(['VDE', 'CATEGORIA', 'SUBCATEGORIA'])['MONTO'].sum().reset_index()
                        
                        # Unimos los datos
                        cumplimiento_subcat = pd.merge(
                            ventas_subcat,
                            presupuesto_subcat,
                            on=['VDE', 'CATEGORIA', 'SUBCATEGORIA'],
                            how='left',
                            suffixes=('_Real', '_Presupuesto')
                        )
                        
                        # Calculamos el cumplimiento
                        cumplimiento_subcat['% Cumplimiento'] = (cumplimiento_subcat['MONTO_Real'] / 
                                                            cumplimiento_subcat['MONTO_Presupuesto']) * 100
                        
                        # Formatear y mostrar
                        st.dataframe(
                            cumplimiento_subcat.style.format({
                                'MONTO_Real': '${:,.2f}',
                                'MONTO_Presupuesto': '${:,.2f}',
                                '% Cumplimiento': '{:.1f}%'
                            }).bar(subset=['% Cumplimiento'], align='mid', color=['#d65f5f', '#5fba7d']),
                            use_container_width=True,
                            height=400
                        )
                    else:
                        st.warning("No se encontraron las columnas necesarias en el presupuesto")
                else:
                    st.warning("No se encontraron las columnas necesarias en los datos de ventas")
            except Exception as e:
                st.error(f"No se pudo calcular el cumplimiento por subcategoría: {str(e)}")

            # Productos por Vendedor (Top 10 - Versión corregida)
            st.markdown("### 📦 Top 10 Productos por Vendedor")
            try:
                # Verificamos las columnas necesarias
                required_cols = ['VDE', 'COD_PROD', 'Descripcion', 'MONTO', 'Cantidad']
                if all(col in current_df.columns for col in required_cols):
                    # Agrupación correcta usando named aggregation
                    top_productos = current_df.groupby(['VDE', 'COD_PROD', 'Descripcion']).agg(
                        Ventas=pd.NamedAgg(column='MONTO', aggfunc='sum'),
                        Cantidad=pd.NamedAgg(column='Cantidad', aggfunc='sum')
                    ).reset_index().sort_values('Ventas', ascending=False)
                    
                    # Mostrar top 10 por vendedor
                    for vendedor in current_df['VDE'].unique():
                        st.markdown(f"#### Vendedor: {vendedor}")
                        df_vendedor = top_productos[top_productos['VDE'] == vendedor].head(60)
                        
                        # Formatear el dataframe para mostrar
                        st.dataframe(
                            df_vendedor[['COD_PROD', 'Descripcion', 'Ventas', 'Cantidad']].style.format({
                                'Ventas': '${:,.2f}',
                                'Cantidad': '{:,.0f}'
                            }),
                            use_container_width=True
                        )
                else:
                    missing_cols = [col for col in required_cols if col not in current_df.columns]
                    st.warning(f"Faltan columnas necesarias: {', '.join(missing_cols)}")
                    st.write("Columnas disponibles:", current_df.columns.tolist())
                    
            except Exception as e:
                st.error(f"No se pudo calcular los productos por vendedor: {str(e)}")
                st.write("Detalle del error:", traceback.format_exc())

    # Módulo de Proyecciones
    elif module == "Proyecciones":
        st.header("📅 Proyecciones para Cierre de Mes")

        # Obtener semana y mes actual
        hoy = datetime.now()
        semana_actual = hoy.isocalendar().week
        mes_actual = hoy.strftime('%B')
        mes_actual_es = MESES_ES.get(mes_actual, mes_actual)

        # Mostrar información actual en la parte superior
        st.subheader(f"📅 Información Actual - Semana {semana_actual}, {mes_actual_es}")

        # Filtros específicos de proyecciones
        with st.sidebar:
            st.header("Filtros de Proyecciones")

            # Obtener solo semanas con datos disponibles
            semanas_con_datos = sorted(sales_df[sales_df['ANO'] == year_filter]['SEM'].unique(), reverse=True)

            # Selección de semana - solo muestra semanas con datos
            semana_filter = st.selectbox(
                "Seleccionar semana",
                options=semanas_con_datos,
                index=0
            )

            # Selección de meses - ahora ordenados cronológicamente
            meses_disponibles = sorted(sales_df['MES'].unique(), key=lambda x: MESES_ORDEN.index(x))
            meses_seleccionados = st.multiselect(
                "Seleccionar mes(es) para proyección",
                options=meses_disponibles,
                default=[mes_actual_es] if mes_actual_es in meses_disponibles else [meses_disponibles[0]] if meses_disponibles else []
            )

            # Selección de vendedores
            vendedores_disponibles = sorted(sales_df['VDE'].unique())
            vendedores_seleccionados = st.multiselect(
                "Seleccionar vendedor(es) para proyección",
                options=vendedores_disponibles,
                default=vendedores_disponibles
            )

        if not meses_seleccionados:
            st.warning("Por favor seleccione al menos un mes para analizar")
            return

        # Filtrar datos por año y vendedores seleccionados
        current_df = sales_df[
            (sales_df['ANO'] == year_filter) &
            (sales_df['VDE'].isin(vendedores_seleccionados))
        ]

        if current_df.empty:
            st.warning("No hay datos disponibles con los filtros seleccionados")
            return

        # Mostrar proyecciones
        col1, col2 = st.columns(2)

        with col1:
            # Proyección Semanal usando el método de la clase KPICalculator
            st.markdown(f"**📅 Proyección Semanal {semana_filter}**")
            proyeccion_semanal = kpi_calculator.calcular_proyeccion_semanal(current_df, semana_filter)

            if proyeccion_semanal:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Ventas Semanales</div>
                    <div class="kpi-value">{format_monto(proyeccion_semanal['ventas_semana'])}</div>
                    <div class="kpi-subtext">Semana {semana_filter}</div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Venta Diaria Promedio</div>
                    <div class="kpi-value">{format_monto(proyeccion_semanal['venta_diaria_promedio'])}</div>
                    <div class="kpi-subtext">{proyeccion_semanal['dias_transcurridos']:.1f} días transcurridos</div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #4CAF50">
                    <div class="kpi-title">Proyección Ajustada</div>
                    <div class="kpi-value">{format_monto(proyeccion_semanal['proyeccion'])}</div>
                    <div class="kpi-subtext">Basado en {proyeccion_semanal['dias_laborables']} días laborables</div>
                </div>
                """, unsafe_allow_html=True)

                # Gráfico evolutivo semanal
                ventas_por_dia = current_df[current_df['SEM'] == semana_filter].groupby('DIA_SEM')['MONTO'].sum().reset_index()

                # Ordenar los días de la semana correctamente
                semana_orden = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                ventas_por_dia['DIA_SEM'] = pd.Categorical(
                    ventas_por_dia['DIA_SEM'],
                    categories=semana_orden,
                    ordered=True
                )
                ventas_por_dia = ventas_por_dia.sort_values('DIA_SEM')

                fig_semana = px.line(
                    ventas_por_dia,
                    x='DIA_SEM',
                    y='MONTO',
                    title=f"Evolución Semana {semana_filter}",
                    markers=True,
                    labels={'MONTO': 'Ventas ($)', 'DIA_SEM': 'Día de la semana'}
                )
                fig_semana.update_layout(
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    xaxis=dict(
                        showgrid=False,
                        categoryorder='array',
                        categoryarray=semana_orden
                    ),
                    yaxis=dict(showgrid=False),
                    height=300,
                    margin=dict(l=40, r=40, t=40, b=40)
                )
                st.plotly_chart(fig_semana, use_container_width=True)
            else:
                st.warning(f"No hay datos para la semana {semana_filter} con los filtros actuales")

        with col2:
            # Proyección Mensual usando el método de la clase KPICalculator
            st.markdown("**📅 Proyección Mensual**")

            if meses_seleccionados:
                mes_actual = meses_seleccionados[0]
                proyeccion_mensual = kpi_calculator.calcular_proyeccion_mensual(current_df, mes_actual)

                if proyeccion_mensual:
                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">Ventas Acumuladas</div>
                        <div class="kpi-value">{format_monto(proyeccion_mensual['ventas_mes'])}</div>
                        <div class="kpi-subtext">Mes {mes_actual}</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">Días Laborables</div>
                        <div class="kpi-value">{proyeccion_mensual['dias_transcurridos']:.1f}/{proyeccion_mensual['dias_totales']:.1f}</div>
                        <div class="kpi-subtext">Días transcurridos/Totales</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">Venta Diaria Promedio</div>
                        <div class="kpi-value">{format_monto(proyeccion_mensual['venta_diaria_promedio'])}</div>
                        <div class="kpi-subtext">Basado en días laborables</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">Ticket Promedio</div>
                        <div class="kpi-value">{format_monto(proyeccion_mensual['ticket_promedio'])}</div>
                        <div class="kpi-subtext">Por transacción</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">Factura Promedio</div>
                        <div class="kpi-value">{format_monto(proyeccion_mensual['factura_promedio'])}</div>
                        <div class="kpi-subtext">Por cliente</div>
                    </div>
                    """, unsafe_allow_html=True)

                    # Gráfico de evolución diaria
                    ventas_diarias = current_df[current_df['MES'] == mes_actual].groupby('Fecha')['MONTO'].sum().reset_index()
                    fig = px.line(
                        ventas_diarias,
                        x='Fecha',
                        y='MONTO',
                        title=f'Evolución de Ventas Diarias - {mes_actual}',
                        markers=True
                    )
                    fig.update_layout(
                        plot_bgcolor='white',
                        paper_bgcolor='white',
                        xaxis=dict(showgrid=False),
                        yaxis=dict(showgrid=False)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning(f"No hay datos para {mes_actual}")

if __name__ == "__main__":
    main()