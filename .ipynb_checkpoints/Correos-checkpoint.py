import win32com.client as win32
from PIL import ImageGrab
#import os
import time
import shutil
from os.path import isfile, join

pathD=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Reportes Industria"

def refreshAndSaveImage(path,fname,path_to_img,images,values):
    xl = win32.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(join(path,fname))
    xl.Visible = True
    wb.RefreshAll()
    shutil.copy(join(path, fname),join(pathD,fname))
    
    for hoja in images.keys():
        ws = wb.Worksheets[hoja]
        for rango in images[hoja].keys():        
            c1=images[hoja][rango][1]
            c2=images[hoja][rango][2]
            ws.Range(ws.Cells(c1[0],c1[1]),ws.Cells(c2[0],c2[1])).CopyPicture(Format = 2)  
            img = ImageGrab.grabclipboard()
            imgFile = join(path_to_img,images[hoja][rango][0])
            img.save(imgFile)
    
    rdict=dict()
    for hoja in values.keys():
        ws = wb.Worksheets[hoja]
        rdict[hoja]=dict()
        for rango in values[hoja].keys():        
            c1=values[hoja][rango][0]
            rdict[hoja][rango]=ws.Cells(c1[0],c1[1]).value
    
    time.sleep(20)
    wb.Close(True)
    time.sleep(20)
    xl.Quit()
    
    if fname=="Ejecución Cuenta 7 (Lite).xlsx":
        xl = win32.DispatchEx("Excel.Application")
        wb = xl.workbooks.open(join(path,"Ejecución Cuenta 7 (Lite) - Completo.xlsx"))
        xl.Visible = True
        wb.RefreshAll()
        shutil.copy(join(path, fname),join(pathD,"Ejecución Cuenta 7 (Lite) - Completo.xlsx"))
        wb.Close(True)
        xl.Quit()
    
    return (rdict)
        
    
    
def correo(html,mailto,subject,path_to_img,images, values, files, tipo, clase):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailto
    mail.Subject = subject
    
    
    if tipo=="Cuenta7":
        contador=0
        for hoja in images.keys():
            for rango in images[hoja].keys():
                att=mail.Attachments.Add(join(path_to_img, images[hoja][rango][0]))
                att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId{}".format(contador))
                contador=contador+1

        for file in files:
            att=mail.Attachments.Add(join(*files[file]))
        
        if clase==1:
            mail.HTMLBody = html.format(values["Resumen STD"]["avanPpto"],
                                    values["Resumen STD"]["varPpto"],
                                    values["Resumen STD"]["cumpPpto"]*100,
                                    "MyId1",
                                    "MyId2",
                                    values["Resumen STD"]["varMesAnt"],
                                    values["Resumen STD"]["varMesMeta"],
                                    "MyId0")
        if clase==2:
            mail.HTMLBody = html
            
        if clase==3:
            mail.HTMLBody = html.format(values["Proyección Variación"]["ckgAct"],
                                    values["Proyección Variación"]["ckgProy"],
                                    values["Proyección Variación"]["mix"],
                                    "MyId0")
            
    if tipo=="Consumos":

        att0=mail.Attachments.Add(join(path_to_img,fnameImage[0]))
        att0.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId0")

        mail.HTMLBody = html.format(varMesAnt,"MyId0")

    if tipo == "BajasDesp":

        att1=mail.Attachments.Add(join(path_to_img, fnameImage[0]))
        att1=mail.Attachments.Add(join(path_to_img, fnameImage[1]))
        mail.HTMLBody = html

    if tipo == "MDR":
        mail.HTMLBody = html
        
    mail.Send()
    
    
def correoC7(clase, test):
    
    if test:
        mailto = 'jcleiva@Grupo-exito.com'
    else:
        mailto = 'jaguirrec@Grupo-Exito.com;yospino@grupo-exito.com;rvillegasa@Grupo-Exito.com;kdelgadillo@Grupo-Exito.com;fbolanos@Grupo-Exito.com;amcorream@grupo-exito.com;scgranadar@grupo-exito.com;jgpenagosp@grupo-exito.com;jherreno@Grupo-Exito.com;jcante@Grupo-Exito.com;xalvarado@Grupo-Exito.com;cploperav@grupo-exito.com;cestepa@Grupo-Exito.com;rvillegasa@Grupo-Exito.com;pparada@grupo-exito.com;lalrodriguezs@Grupo-Exito.com;megarzon@Grupo-Exito.com;jgarciad@Grupo-Exito.com;cfdiaz@Grupo-Exito.com'
    
    if clase==1:
        html="""
        <p>Buen día equipo,<br></p>
        <p>
        El avance en la ejecución de la cuenta 7 es de 
        <b>{:,.0f} MM COP</b>, con una variación frente a presupuesto de 
        <b>{:,.0f} MM COP</span></b> (Cumplimiento: {:,.1f}%)
        </p>
        <p align="center">
        <img src="cid:{}" alt="Tabla Descripción generada automáticamente">
        </p>
        <br>
        <p>El cumplimiento por centro de beneficio (CEBE) es el siguiente:</p>
        <p align="center">
        <img src="cid:{}" alt="Tabla Descripción generada automáticamente">
        </p>
        <br>
        <p>
        La proyección líneal de la variación suponiendo un costo real similar al del mes anterior son
        <b>{:,.0f} MM COP</b>, mientras que versus la meta planteada son
        <b>{:,.0f} MM COP.</b>
        </p>
        <p align="center">
        <img src="cid:{}" alt="Tabla Descripción generada automáticamente"></p>
        <p>
        <br>
        Las visuales frente a presupuesto no contemplan los centros de beneficio de granos ni salas de desposte. El Centro de Costos de Logística OUT se encuentra filtrado también.<br>
        El link del reporte es el siguiente: 
        <a href="https://grupoexito-my.sharepoint.com/:x:/g/personal/jcleiva_grupo-exito_com/ET_eH6v66-1IvZ3_z2MKOo8BDtKp8_Vul_RibHzKbxKM1A"><span class=MsoSmartlink><
        Ejecución Cuenta 7 (Lite).xlsx</span></a></p>

        <p>Cordial saludo,
        <br>JuanL</p>
        </div>"""

        path=r"C:\Users\jcleiva\Documents\Reportes"
        fname="Ejecución Cuenta 7 (Lite).xlsx"
        path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
        images={"Resumen STD":
                    {"R1":["Variación Cuenta 7.jpg",(23,2),(36,11)],
                    "R2":["Variación Ppto.jpg",(8,14),(25,19)],
                    "R3":["Variación por CEBE.jpg",(35,15),(51,19)]}
                   }
        values={"Resumen STD":{
                "varMesAnt":[(36,9)],
                "varMesMeta":[(36,11)],
                "avanPpto":[(25,16)],
                "varPpto":[(25,18)],
                "cumpPpto":[(25,19)],
                }}
        values=refreshAndSaveImage(path,fname,path_to_img,images,values)

        
        subject = 'Informe Costos de Conversión (Cuenta 7)'
    
        files={}
        
        correo(html,mailto,subject,path_to_img,images,values,files,"Cuenta7",clase)
    
    if clase==2:
        html="""
        <p>Buen día equipo,<br></p>
        <p>Adjunto el informe de ejecución de los costos de conversión (Cuenta 7)</p>
        
        <p>Cordial saludo,
        <br>JuanL</p>
        </div>"""

        path=r"C:\Users\jcleiva\Documents\Reportes"
        fname="Ejecución Cuenta 7 (Lite).xlsx"
        path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
        images={}
        values={}
        values=refreshAndSaveImage(path,fname,path_to_img,images,values)


        subject = 'Informe Costos de Conversión (Cuenta 7)'
        files={"File1":[path,fname]}
        images={}
        
        correo(html,mailto,subject,path_to_img,images,values,files,"Cuenta7",clase)
        
    if clase==3:
        html="""
        <p>Buen día equipo,<br></p>
        <p>Adjunto el informe de ejecución de los costos de conversión (Cuenta 7)</p>
        <p>El costo kilo estándar actual y proyectado es {:,.0f}$/Kg y {:,.0f}$/Kg respectivamente, dejando un efecto mix por {:,.0f} MM COP entre el actual y la proyección. Los detalles se presentan a continuación: </p>
        <p align="center">
        <img src="cid:{}" alt="Tabla Descripción generada automáticamente"></p>
        <p>
        El reporte consolidado junto a años anteriores lo pueden encontrar en la ruta: 
        <a href="https://grupoexito-my.sharepoint.com/:f:/g/personal/jcleiva_grupo-exito_com/ErmInbgWwnxFj-KtP0RnzYkBdLB0tuPxMLV8UqkWaCIvYA?e=OAKR8K">
        <span class=MsoSmartlink>
        Reportes Industria</span></a>
        </p>
        <p>Cordial saludo,
        <br>JuanL</p>
        </div>"""

        path=r"C:\Users\jcleiva\Documents\Reportes"
        fname="Ejecución Cuenta 7 (Lite).xlsx"
        path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
        images={"Proyección Variación":
                    {"R1":["Variación Cuenta 7.jpg",(2,2),(12,13)]}}
        values={"Proyección Variación":{
                "ckgAct":[(12,6)],
                "ckgProy":[(12,10)],
                "mix":[(12,13)]
                }}
        values=refreshAndSaveImage(path,fname,path_to_img,images,values)

        subject = 'Informe Costos de Conversión (Cuenta 7)'
        files={"File1":[path,fname]}

        correo(html,mailto,subject,path_to_img,images,values,files,"Cuenta7",clase)


def correoConsumos():
    html= """
    <p>
    Buen día equipo,
    </p>
    <p>A continuación envio el reporte de variaciones en las órdenes de producción, cuya variación actual suma
    <b>{:,.0f} MM COP.</b></p>
    <p align="center"><img src="cid:{}" alt="Tabla Descripción generada automáticamente"></p>
    <p>
    El link del reporte es el siguiente:
    <a href="https://grupoexito-my.sharepoint.com/:x:/g/personal/jcleiva_grupo-exito_com/EWAKvrk8vGxHkpzJWNin06IBMwDw6KwLhErMuFCLvfSeJw?e=YuPqTs"><span class=MsoSmartlink><
    Reporte de Consumos y Precios.xlsx</span></a></p>
    
    <p>Cordial saludo,
    <br>JuanL</p>
    """
    path=r"C:\Users\jcleiva\Documents\Reportes"
    fname="Reporte de Consumos y Precios.xlsx"
    path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
    fnameImage=['Consumos.jpg']
    varMesAnt=refreshAndSaveImage(path,fname,path_to_img,fnameImage,"Consumos")
    
    mailto = 'jcleiva@Grupo-exito.com'
    subject = 'Informe Consumos y Precios'

    correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt/1000000,0,0,0,0,"Consumos",1)
    

def correoBajasDesp():
    html= """
    <p>
    Buen día,
    </p>
    <p>Adjunto envío los reportes de bajas y despachos de Industria.
    </p>
    
    <p>Cordial saludo,
    <br>JuanL</p>
    """
    path=r"C:\Users\jcleiva\Documents\Reportes"
    fname="Reporte de Bajas.xlsx"
    path_to_img=None
    fnameImage=None
    refreshAndSaveImage(path,fname,path_to_img,fnameImage,"BajasDesp")
    
    fname="Informe de Despachos Industria.xlsx"
    refreshAndSaveImage(path,fname,path_to_img,fnameImage,"BajasDesp")
        
    mailto = 'jcleiva@Grupo-exito.com'
    subject = 'Informe de Bajas y Despachos'
    
    correo(html,mailto,subject,path,["Reporte de Bajas.xlsx","Informe de Despachos Industria.xlsx"],0,0,0,0,0,"BajasDesp",1)
    
def correoMDR():
    html= """
    <p>
    Buen día,
    </p>
    <p>Adjunto el Modelo de Rentabilidad de la Industria:
    <a href="https://grupoexito-my.sharepoint.com/:x:/g/personal/jcleiva_grupo-exito_com/EWAKvrk8vGxHkpzJWNin06IBMwDw6KwLhErMuFCLvfSeJw?e=YuPqTs"><span class=MsoSmartlink><
    Reporte de Consumos y Precios.xlsx</span></a></p>
    </p>
    
    <p>Cordial saludo,
    <br>JuanL</p>
    """
    path=r"C:\Users\jcleiva\Documents\Reportes"
    fname="MDR Industria Cad Sum.xlsx"
    path_to_img=None
    fnameImage=None
    refreshAndSaveImage(path,fname,path_to_img,fnameImage,"MDR")
    correo(html,mailto,subject,path,None,0,0,0,0,0,"BajasDesp",1)
    
if __name__ == "__main__":
    correoC7()
    correoConsumos()