import os
import shutil
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
from datetime import datetime
from dotenv import load_dotenv

# Cargar las variables del archivo .env
load_dotenv()

# 1. Configuración de rutas e iniciales
ruta_pendientes = r"C:\Administracion\Comprobantes_Pendientes"
ruta_procesados = r"C:\Administracion\Comprobantes_Procesados"
ruta_excel = r"C:\Administracion\base_proveedores.xlsx"
ruta_plantillas = r"C:\Administracion\plantillas.xlsx" # <--- NUEVO EXCEL

MI_CORREO = os.getenv("MI_CORREO")
MI_PASSWORD = os.getenv("MI_PASSWORD")

if not MI_CORREO or not MI_PASSWORD:
    raise ValueError("⚠️ Error: No se encontraron las credenciales. Verifica tu archivo .env")

# Variables globales para la rotación de mensajes
indice_plantilla_actual = 0

def crear_carpetas_automaticamente():
    try:
        os.makedirs(ruta_pendientes, exist_ok=True)
        os.makedirs(ruta_procesados, exist_ok=True)
        
        df_proveedores = pd.read_excel(ruta_excel)
        
        for proveedor in df_proveedores['Proveedor'].dropna():
            nombre_carpeta = str(proveedor).strip()
            carpeta_destino = os.path.join(ruta_pendientes, nombre_carpeta)
            if not os.path.exists(carpeta_destino):
                os.makedirs(carpeta_destino)
                
    except Exception as e:
        print(f"❌ Error al crear las carpetas desde el Excel: {e}")

def cargar_plantillas():
    """Lee el Excel de plantillas. Si falla, devuelve una por defecto."""
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
        # Si no existe el archivo, usamos una plantilla de respaldo segura
        return [{
            "asunto": "Comprobantes de Pago: {nombre_carpeta} - {fecha_hoy}",
            "cuerpo": """Estimado equipo de {nombre_carpeta},

Adjuntamos {cantidad} comprobante(s) de pago procesado(s) al día de la fecha ({fecha_hoy}).

Ante cualquier consulta, pueden responder a este correo.

Saludos cordiales,
Administración Grupo Yacopini."""
        }]

def procesar_envios():
    global indice_plantilla_actual
    
    try:
        df_proveedores = pd.read_excel(ruta_excel)
    except Exception as e:
        print("❌ Error: No se encuentra la base de proveedores o está abierta.")
        return

    # Cargamos las plantillas frescas en cada revisión
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
                
                correo_destino = filtro.iloc[0]['Email']
                cantidad_pdfs = len(archivos_pdf)
                print(f"➤ Enviando {cantidad_pdfs} comprobante(s) a {nombre_carpeta}...")
                
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                
                # Seleccionar la plantilla actual y rotar
                plantilla = plantillas_correo[indice_plantilla_actual % len(plantillas_correo)]
                indice_plantilla_actual += 1
                
                msg = EmailMessage()
                msg['Subject'] = plantilla["asunto"].format(nombre_carpeta=nombre_carpeta, fecha_hoy=fecha_hoy, cantidad=cantidad_pdfs)
                msg['From'] = MI_CORREO
                msg['To'] = correo_destino
                
                cuerpo = plantilla["cuerpo"].format(nombre_carpeta=nombre_carpeta, fecha_hoy=fecha_hoy, cantidad=cantidad_pdfs)
                msg.set_content(cuerpo)
                
                for archivo in archivos_pdf:
                    with open(os.path.join(ruta_proveedor, archivo), 'rb') as f:
                        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=archivo)
                
                try:
                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.ehlo()
                        smtp.starttls()
                        smtp.login(MI_CORREO, MI_PASSWORD)
                        smtp.send_message(msg)
                    
                    carpeta_destino = os.path.join(ruta_procesados, nombre_carpeta)
                    os.makedirs(carpeta_destino, exist_ok=True)
                    
                    for archivo in archivos_pdf:
                        shutil.move(os.path.join(ruta_proveedor, archivo), os.path.join(carpeta_destino, archivo))
                    
                    print(f"✅ ¡Enviado y archivado correctamente!")
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
print("trabajará solo de 08:00 a 22:00 hs.")
print("Lee plantillas desde: plantillas.xlsx")
print("==================================================\n")

crear_carpetas_automaticamente()

while True:
    ahora = datetime.now()
    hora_actual = ahora.hour

    if 8 <= hora_actual < 22:
        print(f"[{ahora.strftime('%H:%M:%S')}] Revisando carpetas y plantillas...")
        
        crear_carpetas_automaticamente() 
        procesar_envios()
        
        print("Esperando 5 minutos para la próxima revisión...\n")
        time.sleep(300)
    else:
        print(f"[{ahora.strftime('%H:%M:%S')}] Fuera de horario laboral. El bot descansa...")
        time.sleep(1800)