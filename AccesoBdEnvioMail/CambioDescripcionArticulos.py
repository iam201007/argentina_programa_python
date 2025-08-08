import pyodbc
import pandas as pd
import win32com.client as win32

def enviar_mail(destinatarios, destinatariocc, destinatariocco, asunto, mensaje):
    # Aquí iría la lógica para enviar un correo electrónico
    # Por ejemplo, usando smtplib o cualquier otra librería de envío de correos
    print(f"Enviando correo a {destinatarios} con asunto '{asunto}' y mensaje:\n{mensaje}")
    
    # Cuerpo del correo con la tabla HTML
    cuerpo_correo = f"""
    Estimada/o Buen día,<br>
    <br>
    Te pedimos que modifiques la descripción de los siguientes artículos:<br>
    <br>
    {mensaje}
    <br>
    Saludos cordiales,<br>
    Iván A. Marzo<br>
    """

    # Configuración y envío del correo
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Asunto del correo
        mail.Subject = asunto

        # Destinatarios
        mail.To = '; '.join(destinatarios)  # Cambia por la dirección de correo
        mail.CC = destinatariocc   # Opcional
        mail.BCC = destinatariocco   # Opcional

        # Cuerpo del correo en formato HTML
        mail.HTMLBody = cuerpo_correo

        # Enviar el correo
        mail.Send()
        
        print("Correo enviado con éxito.")

    except Exception as e:
        print(f"Ocurrió un error al enviar el correo: {e}")
    
    

# Definir la ruta completa al archivo de la base de datos
db_file_path = "C:BaseDatosVinculos.accdb"  # ¡Cambia esto por tu ruta real!

try:
    # Construir la cadena de conexión
    # Usar el driver adecuado según el tipo de archivo (.mdb o .accdb)
    # y la versión del motor de base de datos que instalaste.

    # Para archivos .accdb (Access 2007-2016)
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_file_path};"
    )

    # Para archivos .mdb (Access 2003 y anteriores)
    # conn_str = (
    #    r"DRIVER={Microsoft Access Driver (*.mdb)};"
    #    f"DBQ={db_file_path};"
    # )

    # Establecer la conexión
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()

    print("Conexión exitosa a la base de datos articulos.")

    # Lista de códigos de artículo que quieres buscar
#   codigos_a_buscar = ('541339', '123456', '789012')
    codigos_a_buscar = ('541339',)

    # Convertir la lista a una cadena de placeholders '?, ?, ?'
    placeholders = ', '.join('?' * len(codigos_a_buscar))

    # Construir y ejecutar la consulta con el operador IN
    query = f"SELECT * FROM ARTICULOS WHERE codart IN ({placeholders})"
    
    # Usar pandas para leer la consulta SQL directamente en un DataFrame
    # Esto es más eficiente que hacerlo de forma manual
    df = pd.read_sql_query(query, cnxn, params=codigos_a_buscar)

    # Imprimir los resultados del DataFrame
    if not df.empty:
        print(f"\n--- Artículos encontrados ({len(df)} en total) ---")
        print(df)
        print(f"\n--- Información del DataFrame ---")
        print(df.info())
        
        # Lista de las columnas que quieres imprimir
        columnas_a_mostrar = ['fecha', 'codart', 'descripcion', 'niv1', 'niv2', 'niv3', 'artqual']
        df1 = df[columnas_a_mostrar]

        # Convertir el dataframe a HTML
        html_table = df1.to_html(index=False) # index=False para no incluir el índice del dataframe

        # Enviar el correo con la tabla HTML
        destinatarios = ["iamarzo@hotmail.com", "iamarzo@hotmail.com"]
        destinatariocc = "iamarzo@hotmail.com"
        destinatariocco = "iamarzo@hotmail.com"
        asunto = "Modificación de descripción de artículos."
        mensaje = html_table
        enviar_mail(destinatarios, destinatariocc, destinatariocco, asunto, mensaje)        

except pyodbc.Error as ex:
    sqlstate = ex.args[0]
    print(f"Error de conexión o consulta: {sqlstate}")

finally:
    # Asegurarse de cerrar la conexión
    if 'cnxn' in locals() and cnxn:
        cnxn.close()
        print("\nConexión cerrada.")