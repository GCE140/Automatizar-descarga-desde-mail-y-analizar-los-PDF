import os, traceback, PyPDF2, re, openpyxl
from datetime import datetime
from imbox import Imbox, messages

# Variables globales
ahora = datetime.now()
formato_fecha = "%d-%m-%Y-%Hh-%Mm-%Ss"
texto_fecha = ahora.strftime(formato_fecha)
dir_base = os.path.abspath(os.getcwd())
nombre_dir_desc = "pdfs_descargados_"+ texto_fecha


def acceso_mail_descarga():

    # 1- Accede al mail y busca todos los mails recibidos de xxxxxx@xxxxxx.com.ar

    host = "imap.gmail.com"
    usuario = "xxxxxx@xxxxxx.com"
    clave = input("\nIngrese la clave del mail xxxxxx@xxxxxx.com: ")
    
    # FALTA QUE DE ERROR SI LA CLAVE ESTÁ MAL
    conexion = Imbox(host, username=usuario, password=clave, ssl=True, ssl_context=None, starttls=False)
    
    os.mkdir(nombre_dir_desc)
    os.chdir(dir_base + "\\" + nombre_dir_desc)
    dir_desc = os.getcwd()

    print("\nChequeando mails nuevos, descargando los archivos si los hay!\n")

    mails_buscados = conexion.messages(sent_from="xxxxxx@xxxxxx.com.ar", unread=True)
    
    for (uid, message) in mails_buscados:
            conexion.mark_seen(uid)
            # marca como vistos

            for idx, adjuntos in enumerate(message.attachments):
                try:
                    att_fn = adjuntos.get("filename")
                    download_path = f"{dir_desc}\{att_fn}"
                    print(download_path)
                    with open(download_path, "wb") as fp:
                        fp.write(adjuntos.get("content").read())
                except:
                    print(traceback.print_exc())
    conexion.logout()


def analizar_pdfs():

    # Cambia a la carpeta de descarga
    os.chdir(dir_base + "\\" + nombre_dir_desc)
    dir_desc = os.getcwd()

    # Recorre el directorio y lista todos los pdf
    extensions = ('.pdf')
    archivos_analizar = []

    for subdir, dirs, files in os.walk(dir_desc):
        for file in files:
            ext = os.path.splitext(file)[-1].lower()
            if ext in extensions:
                archivos_analizar.append(file)

    print()
    
    # 3- Del 1er PDF extraer el CONCEPTO DEL PAGO: AC Factura Compras Na (ej: 0044-00294094)
    # 4- Del 2do PDF extraer el Nro de Certificado (ej: N° 2021044139) y Monto de la Retención (ej: $ 11.242,63)

    if archivos_analizar != []:

        print("Archivos encontrados en la carpeta para analizar:")
        print()
        
        for archivo in archivos_analizar:
            print(archivo)
        print()

        nros_concepto_pago = []
        fechas_vto = []
        nros_certificado = []
        montos_totales = []
        montos_retenidos = []
        fechas_retenciones = []
    
        concepto_pago = re.compile(r"\d\d\d\d-\d\d\d\d\d\d\d\d")
        fecha = re.compile(r"\d\d\/\d\d\/\d\d\d\d")
        nro_certificado = re.compile(r"\d\d\d\d\d\d\d\d\d\d")

        for i in range (len(archivos_analizar)):
            # print(archivos_analizar[i])
            pdf = open(archivos_analizar[i], "rb")
            reader = PyPDF2.PdfFileReader(pdf)
            pag = reader.getPage(0)
            txt = pag.extractText()

            if archivos_analizar[i][0] == "p":
                dato_concepto_pago = concepto_pago.search(txt)
                # print("Concepto de pago: " + dato_concepto_pago.group())
                nros_concepto_pago.append(dato_concepto_pago.group())

                dato_fecha_vto = fecha.search(txt)
                # print("Fecha de vto: " + dato_fecha_vto.group())
                fechas_vto.append(dato_fecha_vto.group())

            elif archivos_analizar[i][0] == "r":
                dato_nro_certificado = nro_certificado.search(txt)
                # print("Nro de certificado: " + dato_nro_certificado.group())
                nros_certificado.append(dato_nro_certificado.group())

                montos = re.findall(r"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})", txt)
                # print("Monto total: " + montos[0])
                montos_totales.append(montos[0])
                # print("Monto retenido: " + montos[1])
                montos_retenidos.append(montos[1])

                dato_fecha_retencion = fecha.search(txt)
                # print("Fecha de retención: " + dato_fecha_retencion.group())
                fechas_retenciones.append(dato_fecha_retencion.group())
            
            pdf.close()
            # print()
        
            # Borrar al final
            # print(nros_concepto_pago)
            # print(fechas_vto)
            # print(nros_certificado)
            # print(montos_totales)
            # print(montos_retenidos)
            # print(fechas_retenciones)
            # print()

            # 5- Crear Excel / Word con todos los datos extraídos. 
            # Col 1 = Factura Nro, Col 2 = Fecha de comprobante, Col 3 = Nro Certif, Col 4 = Monto de la Retención, Col 5 = Fecha retención.
            # Nombrar el Excel Extraxión_datos_"fecha"

            excel = openpyxl.Workbook()
            hoja = excel["Sheet"]

            hoja["A1"] = "Nro de concepto de pago"
            hoja["B1"] = "Fecha de vto"
            hoja["C1"] = "Nro de certificado"
            hoja["D1"] = "Monto total"
            hoja["E1"] = "Monto retenido"
            hoja["F1"] = "Fecha de retención"

            for j in range(len(nros_concepto_pago)):
                hoja["A"+str(j+2)] = nros_concepto_pago[j]

            for k in range(len(fechas_vto)):
                hoja["B"+str(k+2)] = fechas_vto[k]

            for l in range(len(nros_certificado)):
                hoja["C"+str(l+2)] = nros_certificado[l]

            for m in range(len(montos_totales)):
                hoja["D"+str(m+2)] = montos_totales[m]

            for n in range(len(montos_retenidos)):
                hoja["E"+str(n+2)] = montos_retenidos[n]

            for o in range(len(fechas_retenciones)):
                hoja["F"+str(o+2)] = fechas_retenciones[o]

            excel.save("datos_extraidos_el_"+ texto_fecha + ".xlsx")
    
        print("TERMINADO! " + str(len(nros_concepto_pago)*2) + " ARCHIVOS BAJADOS Y ANALIZADOS!\n")
        print("SE CREO EL ARCHIVO datos_extraidos_el_"+ texto_fecha + ".xlsx\n")
        
    elif archivos_analizar == []:
        print("NADA PARA BAJAR Y ANALIZAR\n")
        dir_padre = os.path.dirname(os.getcwd())
        os.chdir(dir_padre)
        os.rmdir(nombre_dir_desc)
        

if __name__ == "__main__":
    acceso_mail_descarga()
    analizar_pdfs()
    os.system('pause')
