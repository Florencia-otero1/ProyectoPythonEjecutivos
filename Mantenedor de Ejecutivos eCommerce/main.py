from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
import os
from dotenv import load_dotenv
import sys
from webdriver_manager.chrome import ChromeDriverManager
from pageObjects.interaccionWeb import *
from pageObjects.descargaArchivo import *
from pageObjects.registroProceso import *
from pageObjects.config import *
from dotenv import load_dotenv
from pageObjects.envioCorreo import EnvioCorreo


def main():
    load_dotenv()
    CONNECTION_STRING = os.getenv('CONNECTION_STRING')
    CONTAINER_NAME = os.getenv('CONTAINER_NAME')
    DIRECTORIO_LOCAL = directorio_local
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT = os.getenv('SMTP_PORT')
    SMTP_USERNAME = os.getenv('SMTP_USERNAME')
    SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')
    descargaArchivo = DescargaArchivo(CONNECTION_STRING, CONTAINER_NAME, DIRECTORIO_LOCAL)
    blob_mas_reciente = descargaArchivo.obtener_blob_mas_reciente()    
    if blob_mas_reciente:
        descargaArchivo.descargar_blob(blob_mas_reciente)
        nombreArchivo = os.path.basename(blob_mas_reciente)
        print(nombreArchivo)  
    

    # INICIALIZAR EL DRIVER
    service = ChromeService(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    driver = webdriver.Chrome(service=service, options=options)

    driver.get(url_qa_login)
    user = os.getenv('USUARIO_LOGIN')
    password = os.getenv('CONTRASEÑA_LOGIN')

    # INGRESAR A LA PAGINA E IR A MANTENEDOR
    registroP = registroProceso(DIRECTORIO_LOCAL, f"Resultado_{nombreArchivo}")
    interaccionW = InteraccionWeb(driver, registroP)
    driver.maximize_window()
    interaccionW.login(user, password)
    driver.get(url_qa_mantenedor)
    time.sleep(5)

    # Crear excel para reporte
    registroP.escribir_encabezados()
    # Ingresar nuevo ejecutivo
    interaccionW.nuevoEjecutivo(f'{DIRECTORIO_LOCAL}/{nombreArchivo}')

    # Mantención de roles
    interaccionW.modificar_roles(f'{DIRECTORIO_LOCAL}/{nombreArchivo}')

    # Dar de baja a los ejecutivo
    interaccionW.baja_usuario(f'{DIRECTORIO_LOCAL}/{nombreArchivo}')
    time.sleep(2)  

    html_body = """
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
        }
        .container {
            margin: 20px;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            background-color: #f9f9f9;
        }
        h1 {
            color: #4CAF50;
        }
        p {
            color: #333;
        }
        .footer {
            margin-top: 20px;
            font-size: 12px;
            color: #777;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>¡Mantención de ejecutivos Completada!</h1>
        <p>La ejecución del <strong>Mantenedor de Ejecutivos</strong> se ha completado exitosamente.</p>
        <p>Adjuntamos el archivo con el resultado para su revisión y seguimiento.</p>
        <p>Gracias.</p>
        <p>Equipo Sonnix.</p>
    </div>
</body>
</html>
"""

    EnvioCorreo.enviar_mail(
    "Resultado mantención de ejecutivos e-commerce",  # subject
    html_body,  # body
    ["nicolas.cristino@bupa.cl", "florencia.otero@bupa.cl"],  # to_emails
    SMTP_SERVER,  # smtp_server
    SMTP_PORT,  # smtp_port
    SMTP_USERNAME,  # smtp_username
    SMTP_PASSWORD,  # smtp_password
    rf'{DIRECTORIO_LOCAL}\Resultado_{nombreArchivo}', # filePath
    html=True  
)
   

    


if __name__ == "__main__":
    main()
