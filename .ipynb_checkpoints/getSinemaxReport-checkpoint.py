from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
from os import listdir
from os.path import isfile, join
import shutil

import argparse
from datetime import datetime, timedelta
import pandas as pd
import warnings
import win32com.client

def month_year_iter( start_month, start_year, end_month, end_year ):
    ym_start= 12*start_year + start_month - 1
    ym_end= 12*end_year + end_month - 1
    for ym in range( ym_start, ym_end ):
        y, m = divmod( ym, 12 )
        yield y, m+1
        
def parse_args():
    parser = argparse.ArgumentParser(description="Esta función determina qué reportes descargar")
    parser.add_argument("-pdl",dest="pdl",action="store_true",help="Descarga histórico de precios")
    parser.add_argument("-desp",dest="desp",action="store_true",help="Descarga Reporte de Despachos")
    parser.add_argument("-ing",dest="ing",action="store_true",help="Descarga Reporte de Ingresos")
    parser.add_argument("-f",dest="fechas",nargs=4,help="Permite indicar de qué período es el reporte, si no se incluye se descargaran los reportes del mes en curso")
    dia=datetime.now()
    parser.set_defaults(pdl=False,desp=False,fechas=[dia.month,dia.year,dia.month+1,dia.year])
    args=parser.parse_args()
    return args

def titlesReport(path,filename,tipo):
    if tipo=="Despachos":
        columns=['Dependencia Despacha', 'Desc. Dependencia Despacha', 'Orden Despacho/Devolucion',
                 'Cod Instalacion', 'Desc. Cod Instalacion', 'Dependencia Recibe', 'Desc. Dependencia Recibe',
                 'Sublinea', 'Desc. Sublinea', 'Plu', 'Desc. Plu', 'Flujo Logístico',       
                 'Planta Industria', 'Desc. Planta Industria', 'Unidades Despachadas',
                 'CtoCantDespachada']
        conv={"Dependencia Despacha":str,"Orden Despacho/Devolucion":str,"Cod Instalacion":str,"Dependencia Recibe":str,"Sublinea":str,"Plu":str}
        rows=1
        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.workbooks.open(join(path,filename))
        xl.Visible = True
        wb.sheets("Despachos y Devoluciones Indust").cells(1,1).Value="Test"
        wb.Close(SaveChanges=1)
        xl.Quit()
        
    if tipo=="PDL":
        columns=['Dependencia',"Desc. Dependencia","Proveedor Plu-Dep","Desc. Proveedor Plu-Dep","Dia",
                 "Estado PluDepHistoria","Desc. Estado PluDepHistoria","$ Precio Venta Historico",
                 "$ CPM Historico","$ Precio Fabrica Historico","$ Costo Neto Historico","$ Costo Sugerido Historico"]
        conv={"Dependencia":str,"Proveedor Plu-Dep":str}
        rows=6
        
    if tipo=="Ingresos":
        columns=["Sublinea","Desc. Sublinea","Clase Marca","Marca","Desc. Marca","Plu","Desc. Plu","Formato",
                 "Desc. Formato","Cadena","Desc. Cadena","Dependencia","Desc. Dependencia","Proveedor","Desc. Proveedor",
                 "# Unidades Totales","$ Ventas sin impuestos","$ Costo"]
        rows=6
        
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
    pd.read_excel(join(path,filename),skiprows=rows,header=None,names=columns,converters=conv).to_excel(join(path,filename),index=None)
    print("Titulos corregidos {}".format(filename))
        
def pdlReport(driver,m,y,mypathD,mypathPDL):
   
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="id_mstr247"]/div[2]/div/div/div[2]/div/a[4]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr264_txt"]').clear()
    driver.find_element(By.XPATH,'//*[@id="id_mstr264_txt"]').send_keys("01/{:02d}/{}".format(m,y))
    driver.find_element(By.XPATH,'//*[@id="id_mstr266"]').click()

    driver.find_element(By.XPATH,'//*[@id="id_mstr253"]').click()

    original_window=driver.current_window_handle
    ###
    time.sleep(50)
    driver.find_element(By.XPATH,'//*[@id="tbExport"]').click()
    time.sleep(3)
    driver.switch_to.window(driver.window_handles[-1])

    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]
    driver.find_element(By.XPATH,'//*[@id="3131"]').click()
    time.sleep(10)
    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. PDL.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    time.sleep(3)
    print("Reporte de pdl descargado: {} {}".format(y,m))
    titlesReport(mypathPDL, "{}{:02d}. PDL.xlsx".format(y,m),"PDL")

def despReport(driver,m,y,mypathD,mypathPDL):

    #time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="id_mstr207"]/div[2]/div/div/div[2]/div/a[4]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr303_txt"]').clear()
    meses={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
          11:"Noviembre",12:"Diciembre"}
    driver.find_element(By.XPATH,'//*[@id="id_mstr303_txt"]').send_keys("{} {}".format(meses[m],y))
    
    driver.find_element(By.XPATH,'//*[@id="id_mstr305"]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr292"]').click()
    original_window=driver.current_window_handle
    
    element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tbExport"]')))
    element.click()
    #time.sleep(50)
    #driver.find_element(By.XPATH,'//*[@id="tbExport"]').click()
    time.sleep(3)

    driver.switch_to.window(driver.window_handles[-1])

    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]
    driver.find_element(By.XPATH,'//*[@id="exportPageByInfo"]').click()
    driver.find_element(By.XPATH,'//*[@id="exportReportTitle"]').click()
    driver.find_element(By.XPATH,'//*[@id="exportFilterDetails"]').click()
    driver.find_element(By.XPATH,'//*[@id="3131"]').click()
    
    time.sleep(80)

    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. Despachos.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    time.sleep(3)
    print("Reporte de Despachos descargado: {} {}".format(y,m))
    titlesReport(mypathPDL, "{}{:02d}. Despachos.xlsx".format(y,m),"Despachos")

def ingReport(driver,m,y,mypathD,mypathPDL):
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="id_mstr116"]/table/tbody/tr[8]/td[2]').click()
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="id_mstr329"]/div[2]/div/div/div[2]/div/a[4]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr460_txt"]').clear()
    meses={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
          11:"Noviembre",12:"Diciembre"}
    driver.find_element(By.XPATH,'//*[@id="id_mstr460_txt"]').send_keys("{} {}".format(meses[m],y))
    driver.find_element(By.XPATH,'//*[@id="id_mstr462"]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr284"]').click()

    original_window=driver.current_window_handle

    time.sleep(50)

    driver.find_element(By.XPATH,'//*[@id="tbExport"]').click()
    time.sleep(3)

    driver.switch_to.window(driver.window_handles[-1])
    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    driver.find_element(By.XPATH,'//*[@id="3131"]').click()

    time.sleep(30)

    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. Ingresos.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    time.sleep(3)
    print("Reporte de Ingresos descargado: {} {}".format(y,m))
    titlesReport(mypathPDL, "{}{:02d}. Ingresos.xlsx".format(y,m),"Ingresos")
    
def getReportSinemax(args):
    
    if args.pdl:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\P&G Industria\V3\Ingresos\PDL"
        service = Service(executable_path=path)
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://pasarela.grupo-exito.com/MicroStrategy/servlet/mstrWeb?evt=3001&src=mstrWeb.3001&Port=0&")

        usr=driver.find_element(By.XPATH,'//*[@id="Uid"]')
        usr.send_keys("jcleiva")
        psw=driver.find_element(By.XPATH,'//*[@id="Pwd"]')
        psw.send_keys("Teb.1030611534")

        driver.find_element(By.XPATH,'//*[@id="3054"]').click()
        driver.find_element(By.XPATH,'//*[@id="projects_ProjectsStyle"]/table/tbody/tr/td[1]/div/table/tbody/tr/td[2]/a').click()
        driver.find_element(By.XPATH,'//*[@id="dktpSectionView"]/a[2]/div[1]').click()
        driver.find_element(By.XPATH,'//*[@id="FolderIcons"]/tbody/tr[2]/td/div/table/tbody/tr/td[2]/a').click()

        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            pdlReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL)
        
    if args.desp:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\Documents\Reportes Base\2023\P&G\Despachos"
        service = Service(executable_path=path)
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://pasarela.grupo-exito.com/MicroStrategy/servlet/mstrWeb?evt=3001&src=mstrWeb.3001&Port=0&")

        usr=driver.find_element(By.XPATH,'//*[@id="Uid"]')
        usr.send_keys("jcleiva")
        psw=driver.find_element(By.XPATH,'//*[@id="Pwd"]')
        psw.send_keys("Teb.1030611534")

        driver.find_element(By.XPATH,'//*[@id="3054"]').click()
        driver.find_element(By.XPATH,'//*[@id="projects_ProjectsStyle"]/table/tbody/tr/td[1]/div/table/tbody/tr/td[2]/a').click()
        driver.find_element(By.XPATH,'//*[@id="dktpSectionView"]/a[2]/div[1]').click()
        driver.find_element(By.XPATH,'//*[@id="FolderIcons"]/tbody/tr[1]/td[2]/div/table/tbody/tr/td[2]/a').click()
        
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            despReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL)
    
    if args.ing:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\P&G Industria\V3\Ingresos\Ingresos"
        service = Service(executable_path=path)
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://pasarela.grupo-exito.com/MicroStrategy/servlet/mstrWeb?evt=3001&src=mstrWeb.3001&Port=0&")

        usr=driver.find_element(By.XPATH,'//*[@id="Uid"]')
        usr.send_keys("jcleiva")
        psw=driver.find_element(By.XPATH,'//*[@id="Pwd"]')
        psw.send_keys("Teb.1030611534")

        driver.find_element(By.XPATH,'//*[@id="3054"]').click()
        driver.find_element(By.XPATH,'//*[@id="projects_ProjectsStyle"]/table/tbody/tr/td[1]/div/table/tbody/tr/td[2]/a').click()
        driver.find_element(By.XPATH,'//*[@id="dktpSectionView"]/a[2]/div[1]').click()
        driver.find_element(By.XPATH,'//*[@id="FolderIcons"]/tbody/tr[1]/td[1]/div/table/tbody/tr/td[2]/a').click()
        
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            ingReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL)
    
    
if __name__ == "__main__":
    args=parse_args()
    getReportSinemax(args)