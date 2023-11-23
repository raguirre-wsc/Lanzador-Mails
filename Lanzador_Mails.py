import win32com.client as win32
import pandas
import arrow

def mailContrapartes():
    #Cramos una instancia de Outlook
    olApp = win32.Dispatch("Outlook.Application")

    #Creamos una lista con Alycs y sus destinatarios
    dest={
        "ADCAP": "fulanito@ad-cap.com.ar; fulanito@ad-cap.com.ar; fulanito@ad-cap.com.ar",
        "ICBC": "fulanito@icbc.com.ar; fulanito@icbc.com.ar",
        "MARIVA": "fulanito@marivafondos.com.ar; fulanito@marivafondos.com.ar; fulanito@mariva.com.ar; fulanito@marivafondos.com.ar",
        "BACS": "fulanito@BACS.COM.AR; fulanito@torontotrust.com.ar",
        "PATAGONIA": "fulanito@bancopatagonia.com.ar; fulanito@bancopatagonia.com.ar; fulanito@bancopatagonia.com.ar; fulanito@bancopatagonia.com.ar; fulanito@bancopatagonia.com.ar",
    }

    #Archivo de trabajo con las operaciones a cursar del dia
    dir_op=r"dir.xlsm"
    
    #Lo levantamos con Pandas y creamos un dataframe
    pd=pandas.read_excel(dir_op, sheet_name="Nueva hoja soporte")

    #Queremos ver las contrapartes (alycs) operadas ese dia. Tomamos los valores unicos del campo "CONTRAPARTE" y convertimos a lista
    contra=list(pd['CONTRAPARTE'].unique())

    #Generamos un loop para crear un mail por cada contraparte
    for i in contra:
        mailItem = olApp.CreateItem(0)

        mailItem.CC = ""
        #El titulo contiene el nombre de la alyc almacenado en "i" y se emplea la libreria arrow para traer la fecha actual
        mailItem.Subject = "SMG - " + i + " - Operaciones del día " + arrow.now().format('DD') + "." + arrow.now().format('MM') + "." + arrow.now().format('YYYY')

        mailItem.BodyFormat = 1
        #Traemos los destinatarios del diccionario creado anteriormente
        mailItem.To = dest[i]

        escritura.mailsContra()
        #Para el cuerpo del mail se utilizara un archivo html creado por la funcion "escritura" que contenga el nombre la alyc
        html = "C:/Users/rodriaguirre/Desktop/Mails/"+i+".html"
        with open(html, "r") as file:
            data = file.read()

        mailItem.HTMLBody = data

        mailItem.Display()

#El objetivo de esta funcion es generar un documento HTML por cada contraparte operada, que contenga una tabla
#con las operaciones a cursar, dandole diferentes formatos al texto en funcion de la operacion detallada.
#Por ej: el campo "MONTO" sera rojo si la operacion es de RESCATE y verde si es una SUSCRIPCION.
def escritura():
    #Archivo de trabajo
    dir_op=r"dir.xlsm"
    #Creamos un df
    pd=pandas.read_excel(dir_op, sheet_name="Nueva hoja soporte")
    #Hacemos manipulaciones a los campos
    pd["MONTO"]=pd["MONTO"].astype('float').round(2)
    pd["OBSERVACIÓN"].fillna("-", inplace=True)
    #Obtenemos una lista con las contrapartes
    contra=list(pd['CONTRAPARTE'].unique())
    #Generamos un html por cada una de ellas
    for i in contra:
        #Filtramos las operaciones aisladas de esa contraparte en un nuevo df
        sub_pd=pd.loc[pd['CONTRAPARTE'] == str(i)].reset_index()
        #Obtenemos la cantidad de operaciones de esa contraparte
        loop=pd.loc[pd['CONTRAPARTE'] == str(i)].shape[0]
        #Generamos el archivo html
        libro = open("C:/Users/rodriaguirre/Desktop/Mails/" + str(i) + ".html", "w")
        html = f"""
        Estimados, les paso las operaciones de fondos para hoy. Por favor, confirmar recepción.<br>
        <br>
        <table>

        <tr>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FECHA OP</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FECHA LIQ</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">OPERACIÓN</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">COMPAÑÍA</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FONDO / TÍTULO</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">MONTO</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">MONEDA</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">OBSERVACIÓN</th>
        </tr>
        """
        #Separamos entre las operaciones de RESCATE y SUSCRIPCION para aplicar diferntes formatos al texto
        for k in range(loop):
            if sub_pd.at[k, "OPERACIÓN"]=="RESCATE":
                html_append = f"""
                <tr>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k, "FECHA OP"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k, "FECHA LIQ"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px; color:red; font-weight: bold">{str(sub_pd.at[k, "OPERACIÓN"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "COMPAÑÍA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "FONDO / TÍTULO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{'{:,.0f}'.format(sub_pd.at[k, "MONTO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "MONEDA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "OBSERVACIÓN"])}</td>
                </tr>
                """
                html += html_append
            else:
                html_append=f"""
                <tr>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k,"FECHA OP"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k,"FECHA LIQ"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"OPERACIÓN"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"COMPAÑÍA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"FONDO / TÍTULO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{'{:,.0f}'.format(sub_pd.at[k, "MONTO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"MONEDA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"OBSERVACIÓN"])}</td>
                </tr>
                """
                html+=html_append


        html_append="""
        </table>
        <br>
        Saludos.
        """
        html += html_append
        #Guardamos el html
        libro.write(html)
