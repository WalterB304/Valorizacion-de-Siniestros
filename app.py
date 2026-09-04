import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, date
import io

st.set_page_config(page_title="Valorización de Seguro de Crédito", layout="wide")

st.title("🏦 Valorización de Créditos Morosos")
st.markdown("Suba el archivo CSV del crédito para calcular automáticamente el valor del siniestro.")

# --- Cálculo de la fecha por defecto (día 10 del mes siguiente) ---
hoy = date.today()
if hoy.month == 12:
    fecha_defecto = date(hoy.year + 1, 1, 10)
else:
    fecha_defecto = date(hoy.year, hoy.month + 1, 10)

# Sidebar - Parámetros de entrada
st.sidebar.header("Parámetros de Entrada")
st.sidebar.text("(*) Valor Obligatorio")

uploaded_file = st.sidebar.file_uploader("(*) Cargar archivo CSV", type=["csv"])

ID_CONSORCIO = "6081e33547d36d680a75ddbb"

if uploaded_file is not None:
    try:

        # Extraer ID por defecto desde el nombre del archivo
        nombre_archivo_raw = uploaded_file.name.rsplit('.', 1)[0]
        partes_nombre = [p.strip() for p in nombre_archivo_raw.split('-')]
        id_operacion_default = partes_nombre[-1] if partes_nombre and len(partes_nombre[-1]) >= 8 else ""

        # Extraer la fecha de otorgamiento por defecto (YYYYMMDD)
        fec_otorga_default = date.today()
        if len(id_operacion_default) >= 8:
            try:
                fec_otorga_default = datetime.strptime(id_operacion_default[:8], "%Y%m%d").date()
            except (ValueError, AttributeError):
                fec_otorga_default = date.today()

        # Datos Sidebar
        tasa_anual = st.sidebar.number_input(
            "(*) Tasa de Interés Anual (Ej. 5,00 para 5,00%)",
            min_value=0.0000,
            max_value=100.00,
            value=0.0,
            step=0.0100,
            format="%.2f"
        )/100

        fecha_valoracion = st.sidebar.date_input(
            "(*) Fecha de Valoración al:",
            value=fecha_defecto,
            format = "DD/MM/YYYY"
        )

        # Leer el CSV
        try:
            # Detectar separador automáticamente (; o ,)
            df = pd.read_csv(uploaded_file, sep=None, engine='python')
        except Exception:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=';', engine='python')

        # Extraer columnas requeridas
        TD = df[['paymentNumber', 'dueDate', 'interest', 'amortization', 'balance', 'isPaid', 'isGraceMonth', 'investorId']].copy()

        # Convertir números si vienen con coma decimal (ej: "7,973994" -> 7.973994)
        for col in ['interest', 'amortization', 'balance']:
            s_clean = TD[col].astype(str).str.replace(',', '.', regex=False).str.strip()
            TD[col] = pd.to_numeric(s_clean, errors='coerce')

        # Convertir fechas (dayfirst=True soporta tanto DD-MM-YYYY como YYYY-MM-DD)
        # Convertir fechas detectando si el formato es ISO (YYYY-MM-DD) o Latino (DD-MM-YYYY)
        s_date_str = TD['dueDate'].astype(str).str.strip()
        sample_date = s_date_str.dropna().iloc[0] if not s_date_str.dropna().empty else ""

        if len(sample_date) >= 4 and sample_date[:4].isdigit():
            # Formato YYYY-MM-DD (Año primero)
            TD['dueDate'] = pd.to_datetime(s_date_str)
        else:
            # Formato DD-MM-YYYY o DD/MM/YYYY (Día primero)
            TD['dueDate'] = pd.to_datetime(s_date_str, dayfirst=True)
        
        # 1. La Cuota es Interés + Amortización
        TD['monthlyPayment'] = TD['interest'] + TD['amortization']
        
        # Columnas calculadas de estructura
        TD['fecha_cuota'] = TD['dueDate'].dt.to_period('M').dt.to_timestamp('M') - pd.offsets.MonthEnd(1)
        
        # 2. Calcular Capital y Saldo en cascada (fila por fila)
        capital_list = []
        saldo_list = []
        
        for idx in range(len(TD)):
            if idx == 0:
                # Primer valor de capital = amortización + saldo (de esa fila)
                cap = TD.loc[idx, 'amortization'] + TD.loc[idx, 'balance']
            else:
                # Desde la segunda fila: capital anterior - amortización anterior
                cap = capital_list[-1] - TD.loc[idx - 1, 'amortization']
            
            capital_list.append(cap)
            saldo_list.append(cap - TD.loc[idx, 'amortization'])
            
        TD['Capital'] = capital_list
        TD['balance'] = saldo_list  # Sobrescribimos la columna original 'balance' con nuestro cálculo en cascada
        
        # Obtener plazo y periodos de gracia
        plazo_credito = int(TD['paymentNumber'].max())
        periodos_gracia = int(df['isGraceMonth'].fillna(False).astype(bool).sum())
        monto_otorgado = TD['Capital'][0]
        
        # Determinar número de cuotas (4 para Consorcio, 6 estándar)
        investor_actual = str(TD['investorId'].dropna().iloc[-1]) if not TD['investorId'].dropna().empty else ""
        n_cuotas = 4 if investor_actual == ID_CONSORCIO else 6

        #Modificar n° de cuotas
        n_cuotas = st.sidebar.number_input(
            "N° de cuotas a Valor Presente",
            min_value= 1,
            max_value= 360,
            value= n_cuotas,
            step= 1
        )
        
        # Buscar primera cuota impaga sin gracia
        idx_inicio = TD[(TD['isPaid'] != True) & (TD['isGraceMonth'] != True)].index.min()
        
        TD['Cuota Actual'] = np.nan
        saldo_insoluto = 0.0
        cuotas_impagas_sum = 0.0
        
        if pd.notna(idx_inicio):
            pos_inicio = TD.index.get_loc(idx_inicio)
            filas_objetivo = TD.index[pos_inicio : pos_inicio + n_cuotas]
            
            # Capitalización diaria
            f_val = pd.Timestamp(fecha_valoracion)
            tasa_diaria = (1 + tasa_anual) ** (1 / 360) - 1
            dias = (f_val - TD.loc[filas_objetivo, 'dueDate']).dt.days
            
            TD.loc[filas_objetivo, 'Cuota Actual'] = (
                TD.loc[filas_objetivo, 'monthlyPayment'] * (1 + tasa_diaria) ** dias
            )
            
            cuotas_impagas_sum = TD.loc[filas_objetivo, 'Cuota Actual'].sum()
            saldo_insoluto = TD.loc[filas_objetivo[-1], 'balance']
        
        total_siniestro = cuotas_impagas_sum + saldo_insoluto
        
        # Mostrar Métricas / Resumen del Siniestro
        st.subheader("📊 Resumen de la Valoración")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tasa Anual", f"{tasa_anual*100:,.2f}%")
        col2.metric("Cuotas Impagas Actualizadas", f"{cuotas_impagas_sum:,.4f}")
        col3.metric("Saldo Insoluto", f"{saldo_insoluto:,.4f}")
        col4.metric("TOTAL SINIESTRO", f"{total_siniestro:,.4f}")
        
        # Formatear la tabla de desarrollo para presentación/exportación
        TD_export = TD.rename(columns={
            'paymentNumber': 'N° Cuota',
            'fecha_cuota': 'Fecha Cuota',
            'dueDate': 'Fecha Venc.',
            'monthlyPayment': 'Cuota',
            'interest': 'Interés',
            'amortization': 'Amortización',
            'balance': 'Saldo',
            'Cuota Actual': 'Cuota Valor Actual'
        })

        TD_export['Fecha Cuota'] = TD_export['Fecha Cuota'].dt.strftime('%d/%m/%Y')
        TD_export['Fecha Venc.'] = TD_export['Fecha Venc.'].dt.strftime('%d/%m/%Y')
        
        cols_orden = ['N° Cuota', 'Fecha Cuota', 'Fecha Venc.', 'Capital', 'Cuota', 'Interés', 'Amortización', 'Saldo', 'Cuota Valor Actual']
        
        st.subheader("📋 Tabla de Desarrollo")
        st.dataframe(TD_export[cols_orden].style.format({
            'Capital': '{:,.4f}',
            'Cuota': '{:,.4f}',
            'Interés': '{:,.4f}',
            'Amortización': '{:,.4f}',
            'Saldo': '{:,.4f}',
            'Cuota Valor Actual': '{:,.4f}'
        }), 
        use_container_width=True,
        hide_index=True
        )

        st.sidebar.header("Datos Cliente")

        id_operacion = st.sidebar.text_input(
            "ID de Operación",
            value=id_operacion_default
        )

        fecha_otorgamiento = st.sidebar.date_input(
            "Fecha de Otorgamiento:",
            value=fec_otorga_default,
            format="DD/MM/YYYY"
        )

        nombre_cliente = st.sidebar.text_input(
            "Nombre Cliente")

        n_RUT = st.sidebar.text_input(
            "RUT")

        inv = st.sidebar.text_input(
            "Inversionista",
            value= 'Consorcio 'if investor_actual == ID_CONSORCIO else ""
        )

        # EXPORTACIÓN A EXCEL
        st.markdown("---")
        st.subheader("📥 Exportar Resultados")
        
        buffer = io.BytesIO()
        
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            
            # 1. Datos del Cliente y Crédito (Bloque Izquierdo - Columnas A y B)
            df_resumen_izq = pd.DataFrame({
                "Parámetro": [
                    "ID de Operación", "Nombre Cliente", "RUT", "Inversionista", 
                    "Fecha Otorgamiento", "Monto Otorgado", "Tasa Emisión", 
                    "Plazo (meses)", "Períodos de Gracia"
                ],
                "Valor": [
                    id_operacion, nombre_cliente, n_RUT, inv, 
                    fecha_otorgamiento.strftime('%d/%m/%Y'), round(capital_list[0], 2), 
                    tasa_anual, plazo_credito, periodos_gracia
                ]
            })
            df_resumen_izq.to_excel(writer, sheet_name='Cálculo', index=False, header=False, startrow=0, startcol=0)
            
            # 2. Resumen de Valoración y Siniestro (Bloque Derecho - Columnas D y E)
            df_resumen_der = pd.DataFrame({
                "Parámetro": [
                    "Valorización al:", "", "Saldo Insoluto", 
                    "Cuotas Impagas Actualiz.", "VALOR SINIESTRO"
                ],
                "Valor": [
                    fecha_valoracion.strftime('%d/%m/%Y'), "", 
                    round(saldo_insoluto, 4), round(cuotas_impagas_sum, 4), 
                    round(total_siniestro, 4)
                ]
            })
            df_resumen_der.to_excel(writer, sheet_name='Cálculo', index=False, header=False, startrow=0, startcol=3)
            
            # 3. Exportar Tabla de Desarrollo (Empieza en la Fila 13)
            TD_excel = TD_export[cols_orden].fillna("")
            TD_excel.to_excel(writer, sheet_name='Cálculo', index=False, startrow=12, startcol=0)

            # Ajustar el ancho de las columnas
            worksheet = writer.sheets['Cálculo']
            worksheet.column_dimensions['A'].width = 22
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['D'].width = 22
            for col in ['C', 'E', 'F', 'G', 'H', 'I']:
                worksheet.column_dimensions[col].width = 15

        # Terminar de construir el archivo
        buffer.seek(0)
        
        # Botón de Descarga
        nombre_archivo = f"Valorizacion_{id_operacion if id_operacion else 'Credito'}.xlsx"
        st.download_button(
            label="Descargar Excel Completo",
            data=buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
else:
    st.info("Por favor, suba un archivo CSV en el panel de la izquierda para comenzar.")
