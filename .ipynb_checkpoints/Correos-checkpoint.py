import win32com.client as win32
from PIL import ImageGrab
import os
import time
import shutil
from os.path import isfile, join

def refreshAndSaveImage(path,fname,path_to_img,fnameImage,tipo):
    xl = win32.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(path+"/"+fname)
    xl.Visible = True
    wb.RefreshAll()
    shutil.copy(join(path, fname),join(r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Reportes Industria",fname))
    if tipo=="Cuenta7":
        ws = wb.Worksheets['Resumen STD']

        ws.Range(ws.Cells(23,2),ws.Cells(36,11)).CopyPicture(Format = 2)  
        img = ImageGrab.grabclipboard()
        imgFile = os.path.join(path_to_img,fnameImage[0])
        img.save(imgFile)
        varMesAnt=ws.Cells(36,9).value
        varMesMeta=ws.Cells(36,11).value
        
        ws.Range(ws.Cells(8,14),ws.Cells(25,19)).CopyPicture(Format = 2)  
        img = ImageGrab.grabclipboard()
        imgFile = os.path.join(path_to_img,fnameImage[1])
        img.save(imgFile)
        avanPpto=ws.Cells(25,16).value
        varPpto=ws.Cells(25,18).value
        cumpPpto=ws.Cells(25,19).value
        
        ws.Range(ws.Cells(35,15),ws.Cells(51,19)).CopyPicture(Format = 2)  
        img = ImageGrab.grabclipboard()
        imgFile = os.path.join(path_to_img,fnameImage[2])
        img.save(imgFile)
        
        time.sleep(20)
        wb.Close(True)
        time.sleep(20)
        xl.Quit()
        return (varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto)
        
    if tipo=="Consumos":
        ws = wb.Worksheets['Resumen']

        ws.Range(ws.Cells(6,2),ws.Cells(16,7)).CopyPicture(Format = 2)  
        img = ImageGrab.grabclipboard()
        imgFile = os.path.join(path_to_img,fnameImage[0])
        img.save(imgFile)
        varMesAnt=ws.Cells(5,7).value
        time.sleep(20)
        wb.Close(True)
        time.sleep(20)
        xl.Quit()
        return (varMesAnt)
    
    if tipo == "BajasDesp":
        wb.Close(True)
        xl.Quit()        
    
    if tipo == "MDR":
        wb.Close(True)
        xl.Quit() 
    
def correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto,tipo, clase):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailto
    mail.Subject = subject
    
    if clase==1:
        if tipo=="Cuenta7":

            att1=mail.Attachments.Add(path_to_img+"/" + fnameImage[1])
            att1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

            att2=mail.Attachments.Add(path_to_img+"/" + fnameImage[2])
            att2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId2")

            att0=mail.Attachments.Add(path_to_img+"/" + fnameImage[0])
            att0.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId0")

            mail.HTMLBody = html.format(avanPpto,varPpto,cumpPpto*100,
                                        "MyId1","MyId2",varMesAnt,varMesMeta, "MyId0")
        if tipo=="Consumos":

            att0=mail.Attachments.Add(path_to_img+"/" + fnameImage[0])
            att0.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId0")

            mail.HTMLBody = html.format(varMesAnt,"MyId0")

        if tipo == "BajasDesp":

            att1=mail.Attachments.Add(path_to_img+"/" + fnameImage[0])
            att1=mail.Attachments.Add(path_to_img+"/" + fnameImage[1])
            mail.HTMLBody = html

        if tipo == "MDR":
            mail.HTMLBody = html
    
    if clase ==2:
        if tipo=="Cuenta7":
            mail.HTMLBody = html
            att1=mail.Attachments.Add(path_to_img+"/" + fnameImage)

        
    mail.Send()
    
    
def correoC7(clase):
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
        fnameImage=['Variación Cuenta 7.jpg','Variación Ppto.jpg','Variación por CEBE.jpg']
        varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto=refreshAndSaveImage(path,fname,path_to_img,fnameImage,"Cuenta7")

        mailto = 'jcleiva@Grupo-exito.com'
        subject = 'Informe Costos de Conversión (Cuenta 7)'

        correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto,"Cuenta7",clase)
    
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
        fnameImage=['Variación Cuenta 7.jpg','Variación Ppto.jpg','Variación por CEBE.jpg']
        varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto=refreshAndSaveImage(path,fname,path_to_img,fnameImage,"Cuenta7")

        mailto = 'jaguirrec@Grupo-Exito.com;yospino@grupo-exito.com;rvillegasa@Grupo-Exito.com;kdelgadillo@Grupo-Exito.com;fbolanos@Grupo-Exito.com;amcorream@grupo-exito.com;scgranadar@grupo-exito.com;jgpenagosp@grupo-exito.com;jherreno@Grupo-Exito.com;jcante@Grupo-Exito.com;xalvarado@Grupo-Exito.com;cploperav@grupo-exito.com;cestepa@Grupo-Exito.com;rvillegasa@Grupo-Exito.com;pparada@grupo-exito.com'
        #mailto = 'jcleiva@Grupo-exito.com'
        subject = 'Informe Costos de Conversión (Cuenta 7)'

        correo(html,mailto,subject,path,fname,varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto,"Cuenta7",clase)

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
    