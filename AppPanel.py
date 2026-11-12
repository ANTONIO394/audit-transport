import os
import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime, timedelta

CARPETA_ENTRADAS = "C:\\AuditTransport\\Albaranes_Entrada"
CARPETA_TARIFAS_DOCS = "C:\\AuditTransport\\Tarifas_Documentos"
EXCEL_MAESTRO = "C:\\AuditTransport\\Master_Expediciones.xlsx"
EXCEL_TARIFAS = "C:\\AuditTransport\\Master_Tarifas.xlsx"
EXCEL_USUARIOS = "C:\\AuditTransport\\Master_Usuarios.xlsx"
CARPETA_HISTORICO = "C:\\AuditTransport\\Historico_Facturas"

CORREO_CONTACTO = "turolpack2000@hotmail.com"
IBAN_OFICIAL = "ES9730800008142707403313"  # Caja Rural de Teruel (Titular: Antonio Jose Martinez Ruperez)
ENLACE_STRIPE = "https://buy.stripe.com/tu_enlace_de_pago"  # Tu enlace de pasarela Stripe

st.set_page_config(page_title="AuditTransportPro - Búnker Operativo", layout="wide")

st.markdown("""
    <style>
    .stButton>button { width: 100%; height: 3em; font-weight: bold; border-radius: 8px; }
    [data-testid="stMetricValue"] { font-size: 1.5rem !important; }
    </style>
""", unsafe_allow_html=True)

for carpeta in [CARPETA_ENTRADAS, CARPETA_TARIFAS_DOCS, CARPETA_HISTORICO]:
    if not os.path.exists(carpeta):
        os.makedirs(carpeta, exist_ok=True)

def inicializar_excel_usuarios():
    if not os.path.exists(EXCEL_USUARIOS):
        df_init = pd.DataFrame(columns=["FechaRegistro", "NombreUsuario", "Correo", "EstadoPrueba", "FechaFinPrueba"])
        df_init.to_excel(EXCEL_USUARIOS, index=False)

inicializar_excel_usuarios()

def leer_excel_seguro(ruta):
    if not os.path.exists(ruta):
        return pd.DataFrame()
    try:
        return pd.read_excel(ruta)
    except PermissionError:
        st.error(f"⚠️ **Error de Bloqueo:** El archivo `{os.path.basename(ruta)}` está abierto en Excel.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"⚠️ Error al leer el fichero: {e}")
        return pd.DataFrame()

def guardar_excel_seguro(df, ruta):
    try:
        df.to_excel(ruta, index=False)
        return True
    except PermissionError:
        st.error(f"❌ **Imposible Guardar:** El archivo `{os.path.basename(ruta)}` está bloqueado.")
        return False
    except Exception as e:
        st.error(f"❌ Error al guardar: {e}")
        return False

# CONTROL DE ACCESO / REGISTRO DE USUARIO (CONTROL DE DESCARGAS)
if "usuario_registrado" not in st.session_state:
    st.session_state["usuario_registrado"] = False

if not st.session_state["usuario_registrado"]:
    st.title("🚚 AuditTransportPro - Control de Acceso")
    st.markdown(f"Por favor, introduce tus datos para inicializar tu entorno seguro. Soporte y consultas: **{CORREO_CONTACTO}**")
    
    with st.form("form_registro"):
        reg_nombre = st.text_input("Nombre o Empresa de Transportes:")
        reg_correo = st.text_input("Correo Electrónico de Contacto:")
        btn_entrar = st.form_submit_button("Acceder al Búnker Operativo")
        
        if btn_entrar:
            if reg_nombre and reg_correo:
                df_usu = leer_excel_seguro(EXCEL_USUARIOS)
                nueva_fila = pd.DataFrame([{
                    "FechaRegistro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "NombreUsuario": reg_nombre,
                    "Correo": reg_correo,
                    "EstadoPrueba": "Activo (3 días)",
                    "FechaFinPrueba": (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")
                }])
                df_usu = pd.concat([df_usu, nueva_fila], ignore_index=True)
                guardar_excel_seguro(df_usu, EXCEL_USUARIOS)
                
                st.session_state["usuario_registrado"] = True
                st.session_state["nombre_usuario"] = reg_nombre
                st.session_state["correo_usuario"] = reg_correo
                st.success("¡Acceso autorizado! Entrando al sistema...")
                st.rerun()
            else:
                st.error("Por favor, rellena tu nombre y correo para continuar.")
    st.stop()

# INTERFAZ PRINCIPAL DEL BÚNKER
emisor_def = st.session_state.get("cfg_emisor", "TUROLPACK S.L. (Entorno Protegido)\nCIF: B00000000\nPolígono Industrial S/N\n44600 Alcañiz (Teruel)")
cliente_def = st.session_state.get("cfg_cliente", "Operador Logístico Genérico S.A.\nCIF: A-00000000\nZaragoza")
forma_pago_def = st.session_state.get("cfg_pago", "Transferencia Bancaria a 60 días fecha factura")
iban_cliente_def = st.session_state.get("cfg_iban", "ES91 0000 0000 0000 0000 0000")

st.title("🚚 AuditTransportPro - Búnker Operativo y Fiscal")
st.markdown(f"Usuario activo: **{st.session_state.get('nombre_usuario', 'Operador')}** | Soporte / Reclamaciones: **{CORREO_CONTACTO}**")

opciones_modulos = [
    "📋 Expediciones", 
    "📤 Subir Albarán",
    "⏱️ Demoras",
    "📄 Facturas PDF",
    "⚙️ Configuración",
    "📥 Subir Tarifas",
    "💶 Editor de Tarifas",
    "📊 Matriz Base", 
    "⛽ Precios Gasoil",
    "👥 Gestión de Descargas",
    "📖 Guía y Ayuda"
]

opcion = st.selectbox("📂 SELECCIONA EL MÓDULO DE TRABAJO:", opciones_modulos)
st.markdown("---")

if opcion == "📋 Expediciones":
    st.subheader("Listado Maestro de Expediciones (Blindado)")
    df_exp = leer_excel_seguro(EXCEL_MAESTRO)
    
    if not df_exp.empty:
        if "DanioMercancia" not in df_exp.columns:
            df_exp["DanioMercancia"] = "No"
        if "ResponsableDanio" not in df_exp.columns:
            df_exp["ResponsableDanio"] = "Ninguno"
        if "Agencia" not in df_exp.columns:
            df_exp["Agencia"] = "Agencia_Alpha"
            
        df_exp["MotivoIncidencia"] = df_exp["MotivoIncidencia"].fillna("")
        df_exp["EstadoEntrega"] = df_exp["EstadoEntrega"].fillna("Pendiente")
        
        total_exp = len(df_exp)
        pendientes = len(df_exp[df_exp["EstadoEntrega"] == "Pendiente"])
        entregados = len(df_exp[df_exp["EstadoEntrega"] == "Entregado"])
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total", total_exp)
        col_m2.metric("Entregados", entregados)
        col_m3.metric("Pendientes", pendientes)
        
        st.markdown("---")
        exp_seleccionada = st.selectbox("Expedición:", df_exp["IdExpedicion"].astype(str).tolist())
        fila_actual = df_exp[df_exp["IdExpedicion"].astype(str) == exp_seleccionada].iloc[0]
        
        nuevo_estado = st.selectbox("Estado Entrega:", ["Pendiente", "Entregado", "No Entregado"], index=["Pendiente", "Entregado", "No Entregado"].index(fila_actual["EstadoEntrega"]) if fila_actual["EstadoEntrega"] in ["Pendiente", "Entregado", "No Entregado"] else 0)
        nuevo_danio = st.selectbox("¿Daños?", ["No", "Sí"], index=["No", "Sí"].index(fila_actual["DanioMercancia"]) if fila_actual["DanioMercancia"] in ["No", "Sí"] else 0)
        nueva_resp = st.selectbox("Responsabilidad:", ["Ninguno", "Origen / Fábrica", "Ruta / Transportista", "Destinatario"], index=0)
        nuevo_motivo = st.text_input("Observaciones:", value=str(fila_actual["MotivoIncidencia"]) if pd.notna(fila_actual["MotivoIncidencia"]) else "")
            
        if st.button("Guardar Cambios Seguros"):
            df_total = leer_excel_seguro(EXCEL_MAESTRO)
            idx = df_total[df_total["IdExpedicion"].astype(str) == exp_seleccionada].index
            if not idx.empty:
                df_total.loc[idx, "EstadoEntrega"] = nuevo_estado
                df_total.loc[idx, "DanioMercancia"] = nuevo_danio
                df_total.loc[idx, "ResponsableDanio"] = nueva_resp
                df_total.loc[idx, "MotivoIncidencia"] = nuevo_motivo
                if guardar_excel_seguro(df_total, EXCEL_MAESTRO):
                    st.success("¡Registro actualizado bajo protocolo de seguridad!")
                    st.rerun()
    else:
        st.info("No hay expediciones registradas.")

elif opcion == "📤 Subir Albarán":
    st.subheader("📤 Motor OCR Local - Subida Segura de Albarán")
    archivo_subido = st.file_uploader("Seleccionar imagen o PDF de albarán", type=["pdf", "jpg", "jpeg", "png"])
    
    if archivo_subido is not None:
        nombre_sano = archivo_subido.name
        ruta_destino = os.path.join(CARPETA_ENTRADAS, nombre_sano)
        try:
            with open(ruta_destino, "wb") as f:
                f.write(archivo_subido.getbuffer())
            st.success(f"¡Albarán '{nombre_sano}' procesado con éxito!")
        except Exception as e:
            st.error(f"Error: {e}")

elif opcion == "⏱️ Demoras":
    st.subheader("⏱️ Control de Tiempos de Espera")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        exp_espera = st.text_input("Referencia Ruta:", "Ruta-Test-01")
        horas_totales = st.number_input("Horas en muelle:", min_value=0.0, value=2.5, step=0.5)
    with col_t2:
        franquicia = st.number_input("Franquicia (Horas):", min_value=0.0, value=1.0, step=0.5)
        tarifa_hora = st.number_input("Tarifa (€/h):", min_value=0.0, value=35.0, step=5.0)
        
    if st.button("Calcular Demora"):
        reclamables = max(0.0, horas_totales - franquicia)
        importe = reclamables * tarifa_hora
        st.session_state["demora_calculada"] = {"exp": exp_espera, "reclamables": reclamables, "importe": importe}
        st.success(f"Demora calculada: {importe:.2f} €")

elif opcion == "📄 Facturas PDF":
    st.subheader("📄 Generación de Facturas PDF (Formato Profesional)")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        num_fac = st.text_input("Nº de Factura Oficial:", "TEST-2026-01")
        cliente_factura = st.text_area("Cliente Destino:", value=cliente_def, height=80)
    with col_f2:
        zona_factura = st.selectbox("Zona Aplicada:", ["B1", "B2", "B3", "B4", "B5"])
        forma_pago_factura = st.text_input("Forma y Plazo de Cobro:", value=forma_pago_def)
        aplicar_iva = st.checkbox("Aplicar IVA oficial (21%)", value=True)
        
    ruta_pdf_potencial = os.path.join(CARPETA_HISTORICO, f"Factura_{num_fac}.pdf")
    
    df_exp_fac = leer_excel_seguro(EXCEL_MAESTRO)
    df_tar_fac = leer_excel_seguro(EXCEL_TARIFAS)
    
    if st.button("Generar Factura Oficial Protegida"):
        lineas_factura = []
        tot_base, tot_iva, tot_total = 0, 0, 0
        
        if not df_exp_fac.empty and not df_tar_fac.empty:
            for _, exp in df_exp_fac.iterrows():
                fecha = str(exp.get("FechaServicio", "2026-02-01"))[:10]
                envio = str(exp.get("IdExpedicion", ""))
                estado_ent = str(exp.get("EstadoEntrega", "Pendiente"))
                danio = str(exp.get("DanioMercancia", "No"))
                resp = str(exp.get("ResponsableDanio", "Ninguno"))
                
                p_real = float(exp.get("KilosRealesDestino", 0)) if pd.notna(exp.get("KilosRealesDestino")) else 0.0
                facturable = float(exp.get("PesoFacturable", p_real)) if pd.notna(exp.get("PesoFacturable")) else p_real
                
                tarifa_fila = df_tar_fac[(df_tar_fac["TramoMin_Kg"] <= facturable) & (df_tar_fac["TramoMax_Kg"] > facturable)]
                if tarifa_fila.empty:
                    tarifa_fila = df_tar_fac.tail(1)
                
                base = float(tarifa_fila[zona_factura].values[0]) if not tarifa_fila.empty and zona_factura in tarifa_fila.columns else 20.0
                iva = round(base * 0.21, 2) if aplicar_iva else 0.0
                total = round(base + iva, 2)
                
                tot_base += base
                tot_iva += iva
                tot_total += total
                
                obs = "OK" if estado_ent == "Entregado" else "PEND."
                if danio == "Sí":
                    obs += f" [DAÑO: {resp}]"
                
                lineas_factura.append({
                    "Fecha": fecha, "Nº Envío": envio,
                    "P. Real (kg)": f"{p_real:,.0f}", "Facturable (kg)": f"{facturable:,.0f}",
                    "Estado / Obs.": obs, "Base (€)": f"{base:.2f} €",
                    "IVA (€)": f"{iva:.2f} €", "Total (€)": f"{total:.2f} €"
                })
        
        extra_demora = st.session_state.get("demora_calculada", {}).get("importe", 0.0)
        if extra_demora > 0:
            iva_dem = round(extra_demora * 0.21, 2) if aplicar_iva else 0.0
            lineas_factura.append({
                "Fecha": "VARIOS", "Nº Envío": "DEMORA", "P. Real (kg)": "-", "Facturable (kg)": "1.5 h",
                "Estado / Obs.": "Espera muelle (>1h franq.)", "Base (€)": f"{extra_demora:.2f} €",
                "IVA (€)": f"{iva_dem:.2f} €", "Total (€)": f"{extra_demora + iva_dem:.2f} €"
            })
            tot_base += extra_demora
            tot_iva += iva_dem
            tot_total += (extra_demora + iva_dem)
            
        df_fac = pd.DataFrame(lineas_factura)
        st.session_state["df_factura_oficial"] = df_fac
        st.session_state["totales_factura"] = {"base": tot_base, "iva": tot_iva, "total": tot_total, "num_fac": num_fac}
        
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        
        doc = SimpleDocTemplate(ruta_pdf_potencial, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        elementos = []
        estilos = getSampleStyleSheet()
        
        estilo_emisor = ParagraphStyle('Emisor', parent=estilos['Normal'], fontSize=9, leading=12, alignment=0, textColor=colors.HexColor("#2D3748"))
        estilo_factura_titulo = ParagraphStyle('FacTit', parent=estilos['Normal'], fontSize=12, leading=14, alignment=2, fontName="Helvetica-Bold", textColor=colors.HexColor("#1A365D"))
        estilo_cliente_titulo = ParagraphStyle('CliTit', parent=estilos['Normal'], fontSize=9.5, leading=12, alignment=0, fontName="Helvetica-Bold", textColor=colors.HexColor("#1A365D"))
        estilo_cliente_datos = ParagraphStyle('CliDat', parent=estilos['Normal'], fontSize=8.5, leading=11, alignment=0, textColor=colors.HexColor("#2D3748"))
        estilo_tot_etq = ParagraphStyle('TotEtq', parent=estilos['Normal'], fontSize=9, alignment=2, fontName="Helvetica-Bold", textColor=colors.HexColor("#1A365D"))
        estilo_tot_val = ParagraphStyle('TotVal', parent=estilos['Normal'], fontSize=10, alignment=2, fontName="Helvetica-Bold", textColor=colors.HexColor("#1A365D"))
        estilo_cobro = ParagraphStyle('Cobro', parent=estilos['Normal'], fontSize=8.5, leading=11, alignment=0, textColor=colors.HexColor("#1A365D"))
        estilo_pie = ParagraphStyle('Pie', parent=estilos['Normal'], fontSize=7.5, leading=10, alignment=1, textColor=colors.HexColor("#718096"))
        
        emisor_texto = emisor_def.replace("\n", "<br/>")
        cliente_texto = cliente_factura.replace("\n", "<br/>")
        
        tabla_cabecera = Table([
            [Paragraph(f"<b>EMISOR:</b><br/>{emisor_texto}", estilo_emisor), 
             Paragraph(f"<b>FACTURA Nº:</b> {num_fac}<br/><b>FECHA:</b> {datetime.now().strftime('%Y-%m-%d')}", estilo_factura_titulo)]
        ], colWidths=[270, 270])
        tabla_cabecera.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 15)]))
        elementos.append(tabla_cabecera)
        elementos.append(Spacer(1, 5))
        
        tabla_cliente = Table([
            [Paragraph("<b>CLIENTE:</b>", estilo_cliente_titulo)],
            [Paragraph(cliente_texto, estilo_cliente_datos)]
        ], colWidths=[540])
        tabla_cliente.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 8), ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ]))
        elementos.append(tabla_cliente)
        elementos.append(Spacer(1, 15))
        
        if not df_fac.empty:
            datos_tabla_pdf = [list(df_fac.columns)] + df_fac.values.tolist()
            t_pdf = Table(datos_tabla_pdf, colWidths=[52, 63, 45, 52, 115, 75, 48, 90])
            t_pdf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 7.5),
                ('BOTTOMPADDING', (0,0), (-1,0), 6), ('TOPPADDING', (0,0), (-1,0), 6),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ]))
            elementos.append(t_pdf)
            elementos.append(Spacer(1, 15))
        
        texto_cobro_pdf = f"<b>CONDICIONES DE COBRO:</b><br/>{forma_pago_factura}<br/><br/><b>CUENTA BANCARIA (IBAN CLIENTE):</b><br/>{iban_cliente_def}"
        t_caja_cobro = Table([[Paragraph(texto_cobro_pdf, estilo_cobro)]], colWidths=[260])
        t_caja_cobro.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EDF2F7")),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 8), ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ]))
        
        datos_totales = [
            [Paragraph("Base Imponible:", estilo_tot_etq), Paragraph(f"{tot_base:,.2f} €", estilo_tot_val)],
            [Paragraph("IVA (21%):", estilo_tot_etq), Paragraph(f"{tot_iva:,.2f} €", estilo_tot_val)],
            [Paragraph("TOTAL FACTURA:", estilo_tot_etq), Paragraph(f"{tot_total:,.2f} €", estilo_tot_val)]
        ]
        t_totales = Table(datos_totales, colWidths=[140, 100])
        t_totales.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EDF2F7")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#1A365D")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 8), ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ]))
        
        t_contenedor_inferior = Table([[t_caja_cobro, t_totales]], colWidths=[280, 260])
        t_contenedor_inferior.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (1,0), (1,0), 'RIGHT')]))
        elementos.append(t_contenedor_inferior)
        
        # FIRMA DISCRETA EN EL PIE DE LA FACTURA
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph("Factura realizada por AuditTransportPro", estilo_pie))
        
        doc.build(elementos)
        st.session_state["ruta_pdf_generado"] = ruta_pdf_potencial
        st.success(f"¡Factura {num_fac} generada con éxito!")

    if "df_factura_oficial" in st.session_state:
        st.markdown("---")
        st.markdown("### 🏢 **Vista Previa de Factura y Descarga**")
        
        # AVISO COMERCIAL Y PASARELA DE PAGO CON AVISO PREVIO A TI
        fecha_fin_prueba = (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")
        st.success(f"🎉 **¡Enhorabuena!** Has generado tu primera factura con éxito. Tu periodo de prueba gratuita está activo hasta el **{fecha_fin_prueba}**. Para continuar usando la herramienta, la cuota es de **10 € al mes**.")
        
        st.markdown("### 💳 **Gestión de Suscripción (10 €/mes):**")
        st.info("Antes de realizar el pago, **avísenos por correo electrónico** para registrar su alta y verificar su entorno operativo:")
        st.markdown(f"📧 **Enviar aviso a Soporte:** `{CORREO_CONTACTO}`")
        
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("⚡ **Opción 1: Pago Online (Stripe)**\n\n👉 [Pagar Suscripción (10 €)](" + ENLACE_STRIPE + ")")
        with col_p2:
            st.warning(f"🏦 **Opción 2: Transferencia / Bizum (Caja Rural de Teruel)**\n\n* **Titular:** Antonio Jose Martinez Ruperez\n* **IBAN:** `{IBAN_OFICIAL}`\n* **Concepto:** Suscripción AuditTransportPro")
            
        st.markdown("---")
        st.dataframe(st.session_state["df_factura_oficial"], width='stretch')
        
        t = st.session_state["totales_factura"]
        st.info(f"💶 **Resumen Contable Oficial:** Base Imponible: **{t['base']:,.2f} €** | IVA (21%): **{t['iva']:,.2f} €** | **TOTAL A PERCIBIR: {t['total']:,.2f} €**")
        
        if "ruta_pdf_generado" in st.session_state and os.path.exists(st.session_state["ruta_pdf_generado"]):
            with open(st.session_state["ruta_pdf_generado"], "rb") as f:
                st.download_button(
                    label="📥 Descargar Factura Oficial PDF",
                    data=f,
                    file_name=f"Factura_{t['num_fac']}.pdf",
                    mime="application/pdf"
                )

elif opcion == "⚙️ Configuración":
    st.subheader("⚙️ Configuración Fiscal y Bancaria")
    nuevo_emisor = st.text_area("Datos Emisor:", value=emisor_def, height=90)
    nuevo_cliente = st.text_area("Cliente por Defecto:", value=cliente_def, height=70)
    nueva_forma_pago = st.text_input("Forma y Plazo de Cobro por Defecto:", value=forma_pago_def)
    nuevo_iban = st.text_input("Cuenta Bancaria del Cliente (IBAN):", value=iban_cliente_def)
    
    st.markdown(f"📧 **Correo Oficial de Soporte y Reclamaciones:** `{CORREO_CONTACTO}` (Fijo por normativa).")
    
    if st.button("Guardar Configuración"):
        st.session_state["cfg_emisor"] = nuevo_emisor
        st.session_state["cfg_cliente"] = nuevo_cliente
        st.session_state["cfg_pago"] = nueva_forma_pago
        st.session_state["cfg_iban"] = nuevo_iban
        st.success("¡Configuración guardada con éxito!")

elif opcion == "📥 Subir Tarifas":
    st.subheader("📥 Subir Tarifas")
    st.file_uploader("Subir documento de tarifas", type=["pdf", "jpg", "png"])

elif opcion == "💶 Editor de Tarifas":
    st.subheader("💶 Editor Visual de Tarifas")
    df_edit = leer_excel_seguro(EXCEL_TARIFAS)
    if not df_edit.empty:
        df_editado = st.data_editor(df_edit, num_rows="dynamic", width='stretch')
        if st.button("💾 Guardar Tarifas"):
            if guardar_excel_seguro(df_editado, EXCEL_TARIFAS):
                st.success("¡Tarifas guardadas!")

elif opcion == "📊 Matriz Base":
    st.subheader("Matriz Base")
    df_matriz = leer_excel_seguro(EXCEL_TARIFAS)
    if not df_matriz.empty:
        st.table(df_matriz)

elif opcion == "⛽ Precios Gasoil":
    st.subheader("⛽ Auditoría y Pérdida Económica por Gasoil")
    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        precio_base = st.number_input("Precio Base (€/l):", value=1.20, step=0.01)
    with col_g2:
        precio_actual = st.number_input("Precio Surtidor (€/l):", value=1.45, step=0.01)
    with col_g3:
        litros_consumidos = st.number_input("Litros consumidos / mes:", value=2500.0, step=100.0)
        
    if st.button("Calcular Pérdida en Euros"):
        dif_litro = precio_actual - precio_base
        perdida_total = dif_litro * litros_consumidos
        st.metric("Pérdida Económica Total", f"{perdida_total:,.2f} €", delta_color="inverse")

elif opcion == "👥 Gestión de Descargas":
    st.subheader("👥 Control de Descargas y Usuarios Registrados")
    df_usu_reg = leer_excel_seguro(EXCEL_USUARIOS)
    if not df_usu_reg.empty:
        total_descargas = len(df_usu_reg)
        st.metric("Total Descargas / Registros Realizados", total_descargas)
        st.markdown("---")
        st.dataframe(df_usu_reg, width='stretch')
    else:
        st.info("Aún no hay registros de usuarios.")

elif opcion == "📖 Guía y Ayuda":
    st.subheader("📖 Guía de Uso de AuditTransportPro")
    st.markdown(f"""
    ### 🧭 Manual Práctico y Contacto:
    * **Soporte y Reclamaciones:** `{CORREO_CONTACTO}`
    * **📋 Expediciones:** Gestión y control de entregas y daños.
    * **📄 Facturas PDF:** Facturas limpias para los clientes del transportista. Al generar la primera, se despliega el aviso comercial indicando que deben avisar primero al correo de soporte antes de efectuar el pago por Stripe o Caja Rural.
    * **👥 Gestión de Descargas:** Panel privado para auditar el número de descargas y quién está usando la aplicación.
    """)
