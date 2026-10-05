import os
import shutil
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
from datetime import datetime
import openpyxl
from dotenv import load_dotenv

# Cargar las variables del archivo .env
load_dotenv()

# 1. Configuración de rutas e iniciales
ruta_pendientes = r"C:\Administracion\Comprobantes_Pendientes"
ruta_procesados = r"C:\Administracion\Comprobantes_Procesados"
ruta_reportes = r"C:\Administracion\Reportes_Diarios"
ruta_excel = r"C:\Administracion\base_proveedores.xlsx"
ruta_plantillas = r"C:\Administracion\plantillas.xlsx"

MI_CORREO = os.getenv("MI_CORREO")
MI_PASSWORD = os.getenv("MI_PASSWORD")
CORREOS_REPORTE = os.getenv("CORREOS_REPORTE")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))

if not MI_CORREO or not MI_PASSWORD:
    raise ValueError("⚠️ Error: No se encontraron las credenciales. Verifica tu archivo .env")

# Variables globales
indice_plantilla_actual = 0
reporte_enviado_hoy = False

def crear_carpetas_automaticamente():
    try:
        os.makedirs(ruta_pendientes, exist_ok=True)
        os.makedirs(ruta_procesados, exist_ok=True)
        os.makedirs(ruta_reportes, exist_ok=True)
        
        df_proveedores = pd.read_excel(ruta_excel)
        
        for proveedor in df_proveedores['Proveedor'].dropna():
            nombre_carpeta = str(proveedor).strip()
            carpeta_destino = os.path.join(ruta_pendientes, nombre_carpeta)
            if not os.path.exists(carpeta_destino):
                os.makedirs(carpeta_destino)
                
    except Exception as e:
        print(f"❌ Error al crear las carpetas desde el Excel: {e}")

def cargar_plantillas():
    try:
        df_plantillas = pd.read_excel(ruta_plantillas)
        plantillas = []
        for index, row in df_plantillas.iterrows():
            plantillas.append({
                "asunto": str(row['Asunto']),
                "cuerpo": str(row['Cuerpo'])
            })
        return plantillas
    except Exception as e:
        return [{
            "asunto": "Comprobantes de Pago: {nombre_carpeta} - {fecha_hoy}",
            "cuerpo": "Estimado equipo de {nombre_carpeta},\n\nAdjuntamos {cantidad} comprobante(s) de pago procesado(s) al día de la fecha ({fecha_hoy}).\n\nSaludos cordiales,\nAdministración Grupo Yacopini."
        }]

def registrar_en_reporte(proveedor, archivo, fecha, hora, cc_destinos):
    """Guarda el registro en un archivo Excel (.xlsx) del día actual."""
    fecha_archivo = datetime.now().strftime("%d_%m_%Y")
    ruta_archivo_excel = os.path.join(ruta_reportes, f"Reporte_Envios_{fecha_archivo}.xlsx")
    
    existe = os.path.exists(ruta_archivo_excel)
    
    if not existe:
        wb = openpyxl.Workbook()
        hoja = wb.active
        hoja.title = "Reporte"
        # Se agrega la nueva columna CC
        hoja.append(['Proveedor', 'Archivo PDF', 'Fecha de Envío', 'Hora de Envío', 'CC'])
    else:
        wb = openpyxl.load_workbook(ruta_archivo_excel)
        hoja = wb.active
        
    # Se guarda el dato de CC en la fila
    hoja.append([proveedor, archivo, fecha, hora, cc_destinos])
    wb.save(ruta_archivo_excel)

def enviar_reporte_diario():
    """Envía el archivo Excel a los correos parametrizados al final del día."""
    fecha_archivo = datetime.now().strftime("%d_%m_%Y")
    ruta_archivo_excel = os.path.join(ruta_reportes, f"Reporte_Envios_{fecha_archivo}.xlsx")
    
    if not os.path.exists(ruta_archivo_excel):
        print("➤ No hubo envíos hoy. No hay reporte que enviar.")
        return
        
    if not CORREOS_REPORTE:
        print("⚠️ Alerta: No se configuró 'CORREOS_REPORTE' en el .env. No se enviará el resumen.")
        return

    print("➤ Generando y enviando reporte diario a supervisores...")
    
    msg = EmailMessage()
    msg['Subject'] = f'Reporte Automático de Comprobantes Enviados - {fecha_archivo.replace("_", "/")}'
    msg['From'] = f"{MI_CORREO}@yacopini.com.ar"
    msg['To'] = CORREOS_REPORTE.split(',')
    
    cuerpo = f"""Estimados,

Adjuntamos el reporte automático detallado en formato Excel con todos los comprobantes de pago procesados y enviados exitosamente durante el día de hoy.

Saludos,
Bot Administrativo - Grupo Yacopini."""
    msg.set_content(cuerpo)
    
    with open(ruta_archivo_excel, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=f"Reporte_Envios_{fecha_archivo}.xlsx")
        
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(MI_CORREO, MI_PASSWORD)
            smtp.send_message(msg)
        print("✅ ¡Reporte diario enviado exitosamente!")
    except Exception as e:
        print(f"❌ Error al enviar el reporte diario: {e}")

def procesar_envios():
    global indice_plantilla_actual
    
    try:
        df_proveedores = pd.read_excel(ruta_excel)
    except Exception as e:
        print("❌ Error: No se encuentra la base de proveedores o está abierta.")
        return

    plantillas_correo = cargar_plantillas()
    archivos_enviados = 0

    for nombre_carpeta in os.listdir(ruta_pendientes):
        ruta_proveedor = os.path.join(ruta_pendientes, nombre_carpeta)
        
        if os.path.isdir(ruta_proveedor):
            archivos_pdf = [f for f in os.listdir(ruta_proveedor) if f.endswith('.pdf')]
            
            if archivos_pdf:
                filtro = df_proveedores[df_proveedores['Proveedor'] == nombre_carpeta]
                
                if filtro.empty:
                    print(f"⚠️ Alerta: El proveedor '{nombre_carpeta}' no está en el Excel.")
                    continue
                
                correo_destino = str(filtro.iloc[0]['Email']).strip()
                
                cc_destinos = ""
                if 'cc' in df_proveedores.columns:
                    valor_cc = filtro.iloc[0]['cc']
                    if pd.notna(valor_cc):  
                        cc_destinos = str(valor_cc).strip()
                
                cantidad_pdfs = len(archivos_pdf)
                print(f"➤ Enviando {cantidad_pdfs} comprobante(s) a {nombre_carpeta}...")
                
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                hora_actual_str = datetime.now().strftime("%H:%M:%S")
                
                plantilla = plantillas_correo[indice_plantilla_actual % len(plantillas_correo)]
                indice_plantilla_actual += 1
                
                msg = EmailMessage()
                msg['Subject'] = plantilla["asunto"].format(nombre_carpeta=nombre_carpeta, fecha_hoy=fecha_hoy, cantidad=cantidad_pdfs)
                msg['From'] = f"{MI_CORREO}@yacopini.com.ar"
                msg['To'] = correo_destino
                
                if cc_destinos:
                    msg['Cc'] = cc_destinos
                
                cuerpo = plantilla["cuerpo"].format(nombre_carpeta=nombre_carpeta, fecha_hoy=fecha_hoy, cantidad=cantidad_pdfs)
                msg.set_content(cuerpo)
                
                for archivo in archivos_pdf:
                    with open(os.path.join(ruta_proveedor, archivo), 'rb') as f:
                        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=archivo)
                
                try:
                    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
                        smtp.login(MI_CORREO, MI_PASSWORD)
                        smtp.send_message(msg)
                    
                    carpeta_destino = os.path.join(ruta_procesados, nombre_carpeta)
                    os.makedirs(carpeta_destino, exist_ok=True)
                    
                    for archivo in archivos_pdf:
                        shutil.move(os.path.join(ruta_proveedor, archivo), os.path.join(carpeta_destino, archivo))
                        # Se agrega cc_destinos al registro final
                        registrar_en_reporte(nombre_carpeta, archivo, fecha_hoy, hora_actual_str, cc_destinos)
                    
                    print(f"✅ ¡Enviado, archivado y registrado correctamente!")
                    archivos_enviados += 1
                    
                    print("⏳ Pausa anti-spam (12s)...")
                    time.sleep(12)
                    
                except Exception as e:
                    print(f"❌ Error al enviar a {nombre_carpeta}: {e}")

# ==========================================
# INTERFAZ Y CICLO PRINCIPAL
# ==========================================
print("==================================================")
print("   BOT DE ENVÍO DE COMPROBANTES - GRUPO YACOPINI  ")
print("==================================================")
print("Hola. Puedes minimizar esta ventana. El sistema")
print("trabajará y registrará todo de 08:00 a 19:00 hs.")
print("==================================================\n")

crear_carpetas_automaticamente()

while True:
    ahora = datetime.now()
    hora_actual = ahora.hour

    if 8 <= hora_actual < 19:
        reporte_enviado_hoy = False 
        
        print(f"[{ahora.strftime('%H:%M:%S')}] Revisando carpetas y plantillas...")
        crear_carpetas_automaticamente() 
        procesar_envios()
        
        print("Esperando 5 minutos para la próxima revisión...\n")
        time.sleep(300)
        
    else:
        if not reporte_enviado_hoy:
            enviar_reporte_diario()
            reporte_enviado_hoy = True
            
        print(f"[{ahora.strftime('%H:%M:%S')}] Fuera de horario laboral. El bot descansa...")
        time.sleep(1800)