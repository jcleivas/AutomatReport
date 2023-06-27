import win32com.client as win32
from PIL import ImageGrab
import os
import time

def refreshAndSaveImage(path,fname,path_to_img,fnameImage,tipo):
    xl = win32.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(path+"/"+fname)
    xl.Visible = True
    wb.RefreshAll()
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

        ws.Range(ws.Cells(6,2),ws.Cells(14,7)).CopyPicture(Format = 2)  
        img = ImageGrab.grabclipboard()
        imgFile = os.path.join(path_to_img,fnameImage[0])
        img.save(imgFile)
        varMesAnt=ws.Cells(14,7).value
        time.sleep(20)
        wb.Close(True)
        time.sleep(20)
        xl.Quit()
        return (varMesAnt)
    
    

def correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto,tipo):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailto
    mail.Subject = subject
    if tipo=="Cuenta7":
        mail.HTMLBody = html.format(avanPpto,varPpto,cumpPpto*100,
                                    path_to_img+"/" + fnameImage[1],path_to_img+"/" + fnameImage[2],
                                    varMesAnt,varMesMeta, path_to_img+"/" + fnameImage[0])
    if tipo=="Consumos":
        mail.HTMLBody = html.format(varMesAnt,path_to_img+"/" + fnameImage[0])
    mail.Send()
    
    
def correoC7():
    html="""<div class="WordSection1">
    <br>
    <br>
    <p class="MsoNormal">Buen día equipo,<br></p>
    <p class="MsoNormal">El avance en la ejecución de la cuenta 7 es de 
    <b><span style="font-size:12.0pt">{:,.0f} MM COP</span></b>, con una variación frente a presupuesto de 
    <b><span style="font-size:12.0pt">{:,.0f} MM COP</span></b> (Cumplimiento: {:,.1f}%)</p>
    <p class="MsoNormal" align="center" style="text-align:center"><img width="759" height="292" style="width:7.9062in;height:3.0416in" id="Imagen_x0020_1" src="{}" alt="Tabla
    Descripción generada automáticamente"><o:p></o:p></p>
    <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>El cumplimiento por centro de beneficio (CEBE) es el siguiente:</p>
    <p class="MsoNormal" align="center" style="text-align:center"><img width="678" height="285" style="width:7.0625in;height:2.9687in" id="Imagen_x0020_2" src="{}" alt="Tabla
    Descripción generada automáticamente"><o:p></o:p></p>
    <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    <p class="MsoNormal">La proyección líneal de la variación suponiendo un costo real similar al del mes anterior son
    <b><span style="font-size:12.0pt">{:,.0f} MM COP</span></b>, mientras que versus la meta planteada son
    <b><span style="font-size:12.0pt">{:,.0f} MM COP</span></b>.<o:p></o:p></p>
    <p class="MsoNormal" align="center" style="text-align:center"><img width="743" height="236" style="width:7.7395in;height:2.4583in" id="Imagen_x0020_3" src="{}" alt="Tabla
    Descripción generada automáticamente"><o:p></o:p></p>
    <p class="MsoNormal" align="center" style="text-align:center"><o:p>&nbsp;</o:p></p>
    <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>
    Las visuales frente a presupuesto no contemplan los centros de beneficio de granos ni salas de desposte. El Centro de Costos de Logística OUT se encuentra filtrado también.<br>
    El link del reporte es el siguiente: 
    <a href="https://grupoexito-my.sharepoint.com/:x:/g/personal/jcleiva_grupo-exito_com/ET_eH6v66-1IvZ3_z2MKOo8BDtKp8_Vul_RibHzKbxKM1A"><span class=MsoSmartlink><
    Ejecución Cuenta 7 (Lite).xlsx</span></a> &nbsp;<o:p></o:p></p>

    <p class="MsoNormal">Cordial saludo,
    <br>JuanL</p>
    </div>"""
    
    path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Reportes Industria"
    fname="Ejecución Cuenta 7 (Lite).xlsx"
    path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
    fnameImage=['Variación Cuenta 7.jpg','Variación Ppto.jpg','Variación por CEBE.jpg']
    varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto=refreshAndSaveImage(path,fname,path_to_img,fnameImage,"Cuenta7")
    
    mailto = 'jcleiva@Grupo-exito.com'
    subject = 'Informe Costos de Conversión (Cuenta 7)'

    correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt,varMesMeta,avanPpto,varPpto,cumpPpto,"Cuenta7")


def correoConsumos():
    html="""<div class="WordSection1">
    <br>
    <br>
    <p class="MsoNormal">Buen día equipo,<br></p>
    <p class="MsoNormal">A continuación envio el reporte de variaciones en las órdenes de producción, cuya variación actual suma
    <b><span style="font-size:12.0pt">{:,.0f} MM COP.</span></b></p>
    <p class="MsoNormal" align="center" style="text-align:center"><img width="838" height="181" style="width:8.7291in;height:1.8854in" id="Imagen_x0020_1" src="{}" alt="Tabla
    Descripción generada automáticamente"><o:p></o:p></p>
    <p class="MsoNormal" align="center" style="text-align:center"><o:p>&nbsp;</o:p></p>
    <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>
    El link del reporte es el siguiente: 
    <a href="https://grupoexito-my.sharepoint.com/:x:/g/personal/jcleiva_grupo-exito_com/EWAKvrk8vGxHkpzJWNin06IBMwDw6KwLhErMuFCLvfSeJw?e=YuPqTs"><span class=MsoSmartlink><
    Reporte de Consumos y Precios.xlsx</span></a> &nbsp;<o:p></o:p></p>

    <p class="MsoNormal">Cordial saludo,
    <br>JuanL</p>
    </div>"""
    
    path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Reportes Industria"
    fname="Reporte de Consumos y Precios.xlsx"
    path_to_img=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes\Imagenes"
    fnameImage=['Consumos.jpg']
    varMesAnt=refreshAndSaveImage(path,fname,path_to_img,fnameImage,"Consumos")
    
    mailto = 'jcleiva@Grupo-exito.com'
    subject = 'Informe Consumos y Precios'

    correo(html,mailto,subject,path_to_img,fnameImage,varMesAnt/1000000,0,0,0,0,"Consumos")
    
    
if __name__ == "__main__":
    correoC7()
    correoConsumos()
    