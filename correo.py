import os
import shutil
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
from datetime import datetime

# 1. Configuración de rutas e iniciales
ruta_pendientes = r"C:\Administracion\Comprobantes_Pendientes"
ruta_procesados = r"C:\Administracion\Comprobantes_Procesados"
ruta_excel = r"C:\Administracion\base_proveedores.xlsx"

MI_CORREO = "karendelpino29@gmail.com"
MI_PASSWORD = "ubxtbsanoaurceuh" # Pon tu nueva clave aquí y no la compartas

# 2. Función principal que hace el trabajo
def procesar_envios():
    try:
        df_proveedores = pd.read_excel(ruta_excel)
    except Exception as e:
        print("❌ Error: No se encuentra el archivo Excel o está abierto. Ciérrelo.")
        return

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
                
                # A partir de aquí todo está correctamente alineado dentro del "if"
                correo_destino = filtro.iloc[0]['Email']
                print(f"➤ Enviando {len(archivos_pdf)} comprobante(s) a {nombre_carpeta}...")
                
                # 1. TEXTO DINÁMICO ANTI-SPAM
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                
                msg = EmailMessage()
                msg['Subject'] = f'Comprobantes de Pago: {nombre_carpeta} - {fecha_hoy}'
                msg['From'] = MI_CORREO
                msg['To'] = correo_destino
                
                # El cuerpo del mensaje ahora cambia según el proveedor y la cantidad de archivos
                cuerpo = f"""Estimado equipo de {nombre_carpeta},

Adjuntamos {len(archivos_pdf)} comprobante(s) de pago procesado(s) al día de la fecha ({fecha_hoy}).

Ante cualquier consulta, pueden responder directamente a este correo.

Saludos cordiales,
Administración Grupo Yacopini."""
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
                    
                    # 2. MICRO-PAUSA ANTI-SPAM
                    print("⏳ Pausa de seguridad anti-spam (12s)...")
                    time.sleep(12)
                    
                except Exception as e:
                    print(f"❌ Error al enviar a {nombre_carpeta}: {e}")
    
    if archivos_enviados == 0:
        print("No hay comprobantes nuevos para enviar.")


# 3. Interfaz intuitiva y ciclo de tiempo
print("==================================================")
print("   BOT DE ENVÍO DE COMPROBANTES - GRUPO YACOPINI  ")
print("==================================================")
print("Hola. Puedes minimizar esta ventana. El sistema")
print("trabajará solo de 08:00 a 19:00 hs.")
print("==================================================\n")

while True:
    ahora = datetime.now()
    hora_actual = ahora.hour

    if 8 <= hora_actual < 22:
        print(f"[{ahora.strftime('%H:%M:%S')}] Revisando carpetas...")
        procesar_envios()
        
        # Pausa de 5 minutos (300 segundos) antes de volver a revisar
        print("Esperando 5 minutos para la próxima revisión...\n")
        time.sleep(300)
    else:
        # Si es de noche, avisa y duerme por 30 minutos antes de chequear la hora de nuevo
        print(f"[{ahora.strftime('%H:%M:%S')}] Fuera de horario laboral. El bot está descansando...")
        time.sleep(1800)