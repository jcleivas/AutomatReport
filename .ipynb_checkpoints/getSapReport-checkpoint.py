import subprocess
import shutil
import time
import win32com.client
import pythoncom
import argparse
from calendar import monthrange
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import Correos



def month_year_iter( start_month, start_year, end_month, end_year ):
    ym_start= 12*start_year + start_month - 1
    ym_end= 12*end_year + end_month - 1
    for ym in range( ym_start, ym_end ):
        y, m = divmod( ym, 12 )
        yield y, m+1

def produccionCMV(x):
    if x["CMv"] in ["101","102"]:
        return -x["Cant_Kgrs"]
    else:
        return x["Cant_Kgrs"]

def cmv(x):
    if x["Clase de movimiento"] in ["101","102"]:
        return "Otras Entregas"
    elif x["Clase de movimiento"] in ["261","262"]:
        return "Consumos"
    elif x["Clase de movimiento"] in ["531","532"]:
        return "Subproductos"
    else:
        return "No encontrado"

def parse_args():
    parser = argparse.ArgumentParser(description="Esta función determina qué reportes descargar")
    parser.add_argument("-c",dest="consumo",action="store_true",help="Descarga reportes de Consumo (MB51, COOISPI)")
    parser.add_argument("-rc",dest="consumoR",action="store_true",help="Ejecuta el reporte de Consumos")
    parser.add_argument("-e",dest="ejec",action="store_true",help="Descarga las ejecuciones")
    parser.add_argument("-eCEBE",dest="ejecCEBE",action="store_true",help="Descarga las ejecuciones CEBE")
    parser.add_argument("-m",dest="maestra",action="store_true",help="Descarga Maestras (MM60)")
    parser.add_argument("-f",dest="fechas",nargs=4,help="Permite indicar de qué período es el reporte, si no se incluye se descargaran los reportes del mes en curso")
    parser.add_argument("-qE",dest="qExcel",action="store_true",help="Cierra la aplicación Excel")
    parser.add_argument("-mB",dest="mb51",action="store_true",help="Descarga MB51")
    parser.add_argument("-mBBajas",dest="mb51B",action="store_true",help="Descarga MB51 Bajas")
    parser.add_argument("-D",dest="despachos",action="store_true",help="Descarga Despachos")
    parser.add_argument("-p",dest="prod",action="store_true",help="Descarga Producción")
    parser.add_argument("-ke24",dest="ke24",action="store_true",help="Descarga ke24")
    dia=datetime.now()
    parser.set_defaults(consumo=False,consumoR=False, ejec=False,ejecCEBE=False,mb51=False,mb51B=False, despachos=False, maestra=False,prod=False,ke24=False,fechas=[dia.month,dia.year,dia.month+1,dia.year])
    args=parser.parse_args()
    return args
    

def sapConnectionBase(cSap):
    try:
        path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(5)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        try:
            connection = application.Children(0)
            i=0
            print("Cerrando {} sesion(es) activa(s)".format(int(connection.children.count)))
            while int(connection.children.count) > 0 and i <5:
                session = connection.Children(0)
                session.findbyid("wnd[0]").close()
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                i=i+1
            
            if cSap:
                return 0
            
            connection = application.OpenConnection("RISE - ERP Produccion")
            time.sleep(5)
            session = connection.Children(0)
            session.findById("wnd[1]").maximize()
            session.findById("wnd[1]/usr/txtRSYST-BNAME").text = "1030611534"
            session.findById("wnd[1]/usr/pwdRSYST-BCODE").text = "Jess.1030611534"
            session.findById("wnd[1]").sendVKey(0)
            
            return session            
            
        except:
            connection = application.OpenConnection("RISE - ERP Produccion")
            time.sleep(5)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()
            session.findById("wnd[1]/usr/txtRSYST-BNAME").text = "1030611534"
            session.findById("wnd[1]/usr/pwdRSYST-BCODE").text = "Jess.1030611534"
            session.findById("wnd[1]").sendVKey(0)
            return session

    except pythoncom.com_error as error:
        hr,msg,exc,arg = error.args
        
        if "The 'Sapgui Component' could not be instantiated." == exc[2]:
            print("No se pudo iniciar SAP, revisa tu conexión a internet/VPN")
        else:
            print(exc[2])
            raise Exception(error)     


def sapConnection(cSap):
    path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(5)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    try:
        connection = application.Children(0)

        i=0
        print("Cerrando {} sesion(es) activa(s)".format(int(connection.children.count)))
        while int(connection.children.count) > 0 and i <5:
            session = connection.Children(0)
            session.findbyid("wnd[0]").close()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            i=i+1
                
    except Exception as e:
        print(e)
        pass
        
    connection = application.OpenConnection("RISE - ERP Produccion")
    session = connection.Children(0)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "1030611534"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Jess.1030611534"
    session.findById("wnd[0]").sendVKey(0)
    return session            


def cooisCabeceras(session,m,y,ruta): #m stands for month, y for year
    i=m
    j=y
    dias=monthrange(j,i)[1]
    diasMesAnt=(datetime(j,i,1)-timedelta(days=1)).day
    fechaIni=datetime(j,i,1)-timedelta(days=1)-timedelta(days=diasMesAnt)+timedelta(days=1)
    mesI=fechaIni.month
    yearI=fechaIni.year
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncoois"
    session.findById("wnd[0]").sendVKey(0)
    #
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    #
    
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = 7
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. Cabeceras de orden.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    
    
def cooisComponentes(session,m,y,ruta): #m stands for month, y for year
    i=m
    j=y
    dias=monthrange(j,i)[1]
    diasMesAnt=(datetime(j,i,1)-timedelta(days=1)).day
    fechaIni=datetime(j,i,1)-timedelta(days=1)-timedelta(days=diasMesAnt)+timedelta(days=1)
    mesI=fechaIni.month
    yearI=fechaIni.year
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncoois"
    session.findById("wnd[0]").sendVKey(0)
    
    #
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    #
    
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = 7
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. Componentes.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    
def cooisAdicionales(session,m,y,ruta): #m stands for month, y for year
    i=m
    j=y
    dias=monthrange(j,i)[1]
    diasMesAnt=(datetime(j,i,1)-timedelta(days=1)).day
    fechaIni=datetime(j,i,1)-timedelta(days=1)-timedelta(days=diasMesAnt)+timedelta(days=1)
    mesI=fechaIni.month
    yearI=fechaIni.year
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncoois"
    session.findById("wnd[0]").sendVKey(0)
    
    #
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOD000"
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    #
    
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = 7
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. Adicionales.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))
    
    
def consumosMB51(session,m,y,ruta):
    i=m
    j=y
    dias=monthrange(j,i)[1]
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    #session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[48]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(5,"BTEXT")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. MB51 (Consumos).xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press() 
    
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))


def bajasMB51(session,m,y,ruta):
    i=m
    j=y
    dias=monthrange(j,i)[1]
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    #session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[48]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(5,"BTEXT")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Bajas".format(j)
    fname="{}. MB51 (Bajas).xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press() 
    
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    
def produccion(session,m,y,ruta):
    i=m
    j=y
    dias=monthrange(j,i)[1]
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzproduccion_kg"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 2
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(2,"LINEAPRODUCCION")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "2"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Producción".format(j)
    fname="{}. Producción.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press() 
    
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))
    
def produccionCarnes(session,m,y,ruta):
    i=m
    j=y
    dias=monthrange(j,i)[1]
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzproduccion_kg"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)
    #session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(1,"LINEAPRODUCCION")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Producción".format(j)
    fname="{}. Producción Carnes.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press() 
    print(1)
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    
def cebeC7(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"Industria":"Cuenta 7 Industria CEBE"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ns_alr_87013326"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
            session.findById("wnd[1]").sendVKey(0)
        except:
            print("")
        
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        
        session.findById("wnd[0]/usr/ctxtPAR_27").text = i
        session.findById("wnd[0]/usr/ctxtPAR_28").text = i
        session.findById("wnd[0]/usr/ctxtPAR_31").text = j
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/tbar[1]/btn[39]").press()
        session.findById("wnd[1]/tbar[0]/btn[2]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(3,"SPRCTR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "1"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Ejecución".format(j)
        fname="{}. {}.xlsx".format(i,grupos[g])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))
    

def despachoKg(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsdr_ent"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSP$00013-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00013-HIGH").text = "{:02d}.{:02d}.{}".format(15,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00010-LOW").text = "7010"
    session.findById("wnd[0]/usr/ctxtSP$00010-HIGH").text = "7900"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(3,"VGBEL")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "3"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = ruta+"\{}\Despachos".format(j)
    
    fname="{}. Despachos 1Q.xlsx".format(i)
    
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[2]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    closeExcel(fname)
    
    print("{} descargado con éxito".format(fname))

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsdr_ent"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSP$00013-LOW").text = "{:02d}.{:02d}.{}".format(16,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00013-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]/usr/ctxtSP$00010-LOW").text = "7010"
    session.findById("wnd[0]/usr/ctxtSP$00010-HIGH").text = "7900"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(3,"VGBEL")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "3"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = ruta+"\{}\Despachos".format(j)
    
    fname="{}. Despachos 2Q.xlsx".format(i)
    
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[2]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    closeExcel(fname)
    
    print("{} descargado con éxito".format(fname))
    
def cebeCV(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"Industria":"Costo de Ventas Industria CEBE"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ns_alr_87013326"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
            session.findById("wnd[1]").sendVKey(0)
        except:
            print("")
        
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        
        session.findById("wnd[0]/usr/ctxtPAR_27").text = i
        session.findById("wnd[0]/usr/ctxtPAR_28").text = i
        session.findById("wnd[0]/usr/ctxtPAR_31").text = j
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/tbar[1]/btn[39]").press()
        session.findById("wnd[1]/tbar[0]/btn[2]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(3,"SPRCTR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "1"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Costo de Ventas".format(j)
        fname="{}. {}.xlsx".format(i,grupos[g])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))
        
def cebeIng(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"Industria":"Ingreso Industria CEBE"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ns_alr_87013326"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
            session.findById("wnd[1]").sendVKey(0)
        except:
            print("")
        
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        
        session.findById("wnd[0]/usr/ctxtPAR_27").text = i
        session.findById("wnd[0]/usr/ctxtPAR_28").text = i
        session.findById("wnd[0]/usr/ctxtPAR_31").text = j
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/tbar[1]/btn[39]").press()
        session.findById("wnd[1]/tbar[0]/btn[2]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(3,"SPRCTR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "1"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Ingresos".format(j)
        fname="{}. {}.xlsx".format(i,grupos[g])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))

def cebeGasto(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"Industria":"Gasto Industria CEBE"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ns_alr_87013326"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
            session.findById("wnd[1]").sendVKey(0)
        except:
            print("")
        
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "1030611534"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        
        session.findById("wnd[0]/usr/ctxtPAR_27").text = i
        session.findById("wnd[0]/usr/ctxtPAR_28").text = i
        session.findById("wnd[0]/usr/ctxtPAR_31").text = j
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/tbar[1]/btn[39]").press()
        session.findById("wnd[1]/tbar[0]/btn[2]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(3,"SPRCTR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "1"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Gasto".format(j)
        fname="{}. {}.xlsx".format(i,grupos[g])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))

    
def ksb1(session,m,y,ruta):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"Industria":"Cuenta 7 Industria"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
            session.findById("wnd[1]").sendVKey(0)
        except:
            session.findById("wnd[0]/usr/ctxtP_KOKRS").text = "CO10"
        
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        
        session.findById("wnd[0]/usr/ctxtKSTGR").text = ""
        
        session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
        session.findById("wnd[0]/usr/btnBUT1").press()
        session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(1,"OBJ_TXT")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "1"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Ejecución".format(j)
        fname="{}. {}.xlsx".format(i,grupos[g])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))
        
        print('Cierre los exceles correspondientes')
        x = input()
        print('Exceles cerrados')
        
        ejecAgg(m,y,ruta)

def ejecAgg(m,y,ruta):
    colsEjec=["Clase de coste","Denom.clase de coste","Centro de coste","Denominación del objeto",
          "Material","Texto breve de material","Texto de cabecera de documento",
          "Valor/mon.inf.","Desv.precio fija","Cantidad total"]
    convEjec={"Clase de coste":str,"Centro de coste":str,"Material":str,
              "Valor/mon.inf.":float,"Desv.precio fija":float,"Cantidad total":float,
             "Texto de cabecera de documento":str}
    tiempo=(y,m)
    dfEj=pd.read_excel(ruta+"\{}\Ejecución\{}. Cuenta 7 Industria.xlsx".format(tiempo[0],tiempo[1]),usecols=colsEjec,converters=convEjec)
    
    for i in dfEj.columns:
        if dfEj[i].dtype == "object":
            dfEj[i].fillna("",inplace=True)
        elif (dfEj[i].dtype == "float64") or (dfEj[i].dtype == "int64"):
            dfEj[i].fillna(0.0,inplace=True)

    dfEj["Distribución"]=dfEj["Texto de cabecera de documento"].apply(lambda x: x if x[:2]=="DR" else "")
    del dfEj["Texto de cabecera de documento"]

    colsTemp=["Clase de coste","Denom.clase de coste","Centro de coste","Denominación del objeto",
              "Material","Texto breve de material","Distribución"]
    dfEj=dfEj.groupby(colsTemp,dropna=False).sum().reset_index()
    dfEj["Fecha"]=datetime(tiempo[0],tiempo[1],1)
    
    dfEj=dfEj.rename(columns={"Clase de coste":"Cuenta","Denom.clase de coste":"Denominación Cuenta",
                        "Denominación del objeto":"Denominación Centro de Costo",
                        "Valor/mon.inf.":"Valor Real","Cantidad total":"Horas Reales",
                        "Desv.precio fija":"Variación"})
    dfEj.to_excel(ruta+"\{}\Ejecución\{}. Cuenta 7 Industria (Agg).xlsx".format(tiempo[0],tiempo[1]),index=None)

    
    try:
        colsPpto=["Clase de coste","Denom.clase de coste","Centro de coste","Denominación del objeto","Valor/mon.inf.",
          "Cantidad total","Material","Texto breve de material","Distribución"]
        dfPpto=pd.read_excel(ruta+"\{}\Ejecución\{}. Cuenta 7 Industria Ppto.xlsx".format(tiempo[0],tiempo[1]),usecols=colsPpto,
                             converters={"Clase de coste":str,"Centro de coste":str})

        dfPpto=dfPpto.rename(columns={"Clase de coste":"Cuenta","Denom.clase de coste":"Denominación Cuenta",
                                     "Denominación del objeto":"Denominación Centro de Costo",
                                     "Cantidad total":"Horas Reales","Valor/mon.inf.":"Valor Ppto",})

        dfPpto["Valor Real"]=0
        dfPpto["Variación"]=0
        dfPpto["Fecha"]=datetime(tiempo[0],tiempo[1],1)
        dfPpto["Tipo"]="Ppto"

        dfEj["Valor Ppto"]=0
        dfEj["Tipo"]="Real"

        dfEj=pd.concat([dfEj,dfPpto],sort=False)

    except FileNotFoundError:
        dfEj["Valor Ppto"]=0
        dfEj["Tipo"]="Real"


    dfEj["Centro de beneficio"]=dfEj["Centro de coste"].str[:4]

    dfM=pd.read_excel(rutaM+"\Maestro CEBE Industria.xlsx",usecols=["Centro de beneficio","Denominación CEBE"],converters={"Centro de beneficio":str})

    if dfEj.merge(dfM,on=["Centro de beneficio"],how="left").shape[0]!=dfEj.shape[0]:
        raise Exception("Maestra CEBES inserta filas")

    dfEj=dfEj.merge(dfM,on=["Centro de beneficio"],how="left")

    dfM=pd.read_excel(rutaM+"\Maestra Cuentas.xlsx",
                      usecols=["Clcoste","TIPO PYG_2","TIPO 2 P&G","NOMBREGRUPO PYG"],
                      converters={"Clcoste":str})

    dfM=dfM.rename(columns={"Clcoste":"Cuenta","TIPO PYG_2":"Tipo P&G","NOMBREGRUPO PYG":"Grupo P&G","TIPO 2 P&G":"Tipo P&G 2"})
    if dfEj.merge(dfM,on=["Cuenta"],how="left").shape[0]!=dfEj.shape[0]:
        raise Exception("Maestra Cuentas inserta filas")

    dfEj=dfEj.merge(dfM,on=["Cuenta"],how="left")

    dfM=pd.read_excel(rutaM+"\Centro-CEBE.xlsx",
                      usecols=["Centro","Planta","Cebe"],
                      converters={"Centro":str,"Cebe":str})

    dfM=dfM.rename(columns={"Planta":"Descripción Centro","Cebe":"Centro de beneficio"})
    if dfEj.merge(dfM,on=["Centro de beneficio"],how="left").shape[0]!=dfEj.shape[0]:
        raise Exception("Maestra Centros inserta filas")

    dfEj=dfEj.merge(dfM,on=["Centro de beneficio"],how="left")
    
    dfProd=pd.read_excel(ruta+"\{}\Producción\{}. Producción Carnes.xlsx".format(tiempo[0],tiempo[1]),
                     usecols=["Ce.","Cant_Kgrs","CMv"],
                    converters={"Ce.":str,"CMv":str})

    dfProdAux=pd.read_excel(ruta+"\{}\Producción\{}. Producción.xlsx".format(tiempo[0],tiempo[1]),
                         usecols=["Ce.","Cant_Kgrs","CMv"],
                        converters={"Ce.":str,"CMv":str})
    dfProd=pd.concat([dfProd,dfProdAux],sort=False)
    del dfProdAux

    dfProd=dfProd.rename(columns={"Ce.":"Centro"})

    dfProd["Cantidades"]=dfProd.apply(produccionCMV,axis=1)

    dfProd=dfProd[["Centro","Cantidades"]].groupby(["Centro"]).sum().reset_index()

    d={"Centro":["7300","7310","7736","7737","7738","7743","7808"],
      "Cantidades":[0.0]*len(["7300","7310","7736","7737","7738","7743","7808"])}

    dfProd=pd.concat([dfProd,pd.DataFrame.from_dict(d)],sort=False).groupby(["Centro"]).sum().reset_index()

    dfEj=dfEj.merge(dfProd,on=["Centro"],how="left")

    dfEj["Cantidades"].fillna(dfProd.loc[~dfProd["Centro"].isin(["7300","7310"]),"Cantidades"].sum(),inplace=True)
    dfEj["CtoKg"]=dfEj["Valor Real"].divide(dfEj["Cantidades"])
    
    dfEj["Cantidades2"]=dfProd.loc[~dfProd["Centro"].isin(["7300","7310"]),"Cantidades"].sum()
    dfEj.loc[dfEj["Centro"].isin(["7300","7310"]),"Cantidades2"]=dfEj.loc[dfEj["Centro"].isin(["7300","7310"]),"Cantidades"]
    dfEj["CtoKg2"]=dfEj["Valor Real"].divide(dfEj["Cantidades2"])
    dfEj["CtoKg"].replace([np.inf, -np.inf], 0, inplace=True)
    dfEj["CtoKg2"].replace([np.inf, -np.inf], 0, inplace=True)
    
    dfEj.to_excel(ruta+"\{}\Ejecución\{}. Cuenta 7 Industria (Agg Lite).xlsx".format(tiempo[0],tiempo[1]),index=None)
    
    print("Ejecución Agregada generada con éxito")
    
def maestras(session,ruta):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm60"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(1,"KTEXT")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\Maestras"
    fname="MM60.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname

    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)
    print("{} descargado con éxito".format(fname))


def ke24(session,m,y,ruta):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nke24"
    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById("wnd[1]/usr/radRKEA2-PA_TYPE_2").select()
        session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "pa10"
        session.findById("wnd[1]/usr/radRKEA2-PA_TYPE_2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    session.findById("wnd[0]/mbar/menu[1]/menu[0]/menu[0]").select()
    
    
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn("VARIANT")
    session.findById("wnd[1]/tbar[0]/btn[29]").press()
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Ingresos"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/txtPERIO-LOW").text = "{:03d}.{}".format(m,y)
    session.findById("wnd[0]/usr/txtPERIO-HIGH").text = "{:03d}.{}".format(m,y)
    session.findById("wnd[0]/usr/ctxtHZDAT-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtHZDAT-HIGH").text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(1,"PRCTR")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "14"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Sublineas".format(y)
    fname="{}. Ingresos.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nke24"
    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById("wnd[1]/usr/radRKEA2-PA_TYPE_2").select()
        session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "pa10"
        session.findById("wnd[1]/usr/radRKEA2-PA_TYPE_2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "CO10"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    session.findById("wnd[0]/mbar/menu[1]/menu[0]/menu[0]").select()

    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn("VARIANT")
    session.findById("wnd[1]/tbar[0]/btn[29]").press()
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "COSTOVENTAS"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
    
    session.findById("wnd[0]/usr/txtPERIO-LOW").text = "{:03d}.{}".format(m,y)
    session.findById("wnd[0]/usr/txtPERIO-HIGH").text = "{:03d}.{}".format(m,y)
    session.findById("wnd[0]/usr/ctxtHZDAT-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtHZDAT-HIGH").text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(1,"PRCTR")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "14"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Sublineas".format(y)
    fname="{}. Costo de Ventas.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

    

   
rutaM=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes Base\Maestras"
rutaR=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes"    
ruta=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Reportes Base"

def reporteConsumos(m,y,ruta,rutaM=rutaM,rutaR=rutaR):
    tiempo=(y,m)
    
    colsComp=["Orden","Centro","Material","Texto breve material",
          "Unidad medida base (=EINHEIT)","Cantidad necesaria (EINHEIT)",
          "Cantidad tomada (EINHEIT)","Valor de la toma (WAERS)","Clase de movimiento",
          "Precio/MonL (WAERS)"]

    convComp={"Orden":str,"Centro":str,"Material":str,
              "Clase de movimiento":str}
    
    colsCab=["Centro","Orden","Número material","Texto breve material",
         "Cantidad orden (GMEIN)","Cantidad entregada (GMEIN)",
         "Unidad de medida (=GMEIN)","Status de sistema","Fecha liberac.real"]

    convCab={"Centro":str,"Orden":str,"Número material":str,"Difer.confirm.proc.":str,"División":str}
    
    colsAdi=["Orden","Centro","Material","Texto de material",
         "Unidad medida base (=MEINS)","Ctd.en UM base (MEINS)","Importe ML (WAERS)",
         "Clase de movimiento"]

    convAdi={"Orden":str,"Centro":str,"Material":str,"Clase de movimiento":str}

    colsMb51=["Centro","Orden","Material","Ctd.en UM entrada","Importe ML","Clase de movimiento"]
    convMb51={"Centro":str,"Orden":str,"Material":str,"Clase de movimiento":str}
    
    colsEjec=["Clase de coste","Denom.clase de coste","Orden partner","Valor/mon.inf.","Cantidad total","Desv.precio fija"]
    convEjec={"Clase de coste":str,"Orden partner":str}
    
    dfComp=pd.read_excel(ruta+"\{}\Consumos\{}. Componentes.xlsx".format(tiempo[0],tiempo[1]),usecols=colsComp,converters=convComp)
    dfAdi=pd.read_excel(ruta+"\{}\Consumos\{}. Adicionales.xlsx".format(tiempo[0],tiempo[1]),usecols=colsAdi,converters=convAdi)
    dfCab=pd.read_excel(ruta+"\{}\Consumos\{}. Cabeceras de orden.xlsx".format(tiempo[0],tiempo[1]),usecols=colsCab,converters=convCab)
    
    dfAdi=dfAdi.rename(columns={"Unidad medida base (=MEINS)":"Unidad medida base (=EINHEIT)",
                     'Ctd.en UM base (MEINS)':'Cantidad tomada (EINHEIT)',
                     "Importe ML (WAERS)":'Valor de la toma (WAERS)',
                    "Texto de material":"Texto breve material"})
    
    dfAdi["Cantidad necesaria (EINHEIT)"]=0
    dfAdi["Precio/MonL (WAERS)"]=dfAdi["Valor de la toma (WAERS)"].divide(dfAdi["Cantidad tomada (EINHEIT)"],fill_value=0)
    dfAdi["Precio/MonL (WAERS)"].replace([np.inf, -np.inf], 0, inplace=True)
    dfComp=pd.concat([dfComp,dfAdi],sort=False)
    del dfAdi

    dfComp["Clase"]=dfComp.apply(cmv,axis=1)
    dfCab["Status"]=dfCab["Status de sistema"].str[:4]

    dfCab=dfCab.rename(columns={"Cantidad orden (GMEIN)":"Cantidad Plan PT","Cantidad entregada (GMEIN)":"Cantidad Real PT",
                               'Número material':"Receta",'Texto breve material':"Desc. Receta"})
    
    temp=dfCab[["Centro","Orden","Receta"]].copy()
    temp=temp.rename(columns={"Receta":"Material"})
    temp["Cab"]="Cab"
    temp["Clase de movimiento"]="101"

    temp2=temp.copy()
    temp2["Clase de movimiento"]="102"
    temp=pd.concat([temp,temp2])
    del temp2

    if dfComp.merge(temp,on=["Centro","Orden","Material","Clase de movimiento"],how="left").shape[0]!=dfComp.shape[0]:
        raise Exception("Temp inserta columnas")

    dfComp=dfComp.merge(temp,on=["Centro","Orden","Material","Clase de movimiento"],how="left")

    dfComp=dfComp[dfComp["Cab"].isna()]
    del dfComp["Cab"]

    dfMb51=pd.read_excel(ruta+"\{}\Consumos\{}. MB51 (Consumos).xlsx".format(tiempo[0],tiempo[1]),usecols=colsMb51+["Unidad medida base"],converters=convMb51)
    dfMb51=dfMb51.groupby(["Orden","Centro","Material","Clase de movimiento","Unidad medida base"]).sum().reset_index()
    if dfMb51.merge(temp,on=["Centro","Orden","Material","Clase de movimiento"],how="left").shape[0]!=dfMb51.shape[0]:
        raise Exception("Temp inserta columnas")

    dfMb51=dfMb51.merge(temp,on=["Centro","Orden","Material","Clase de movimiento"],how="left")
    dfMb51=dfMb51[~dfMb51["Cab"].isna()]
    del dfMb51["Cab"]

    temp=dfCab.copy()

    temp=temp.rename(columns={"Receta":"Material","Desc. Receta":'Texto breve material',
                             "Unidad de medida (=GMEIN)":'Unidad medida base (=EINHEIT)',
                             "Cantidad Plan PT":'Cantidad necesaria (EINHEIT)',
                              "Cantidad Real PT":'Cantidad tomada (EINHEIT)'})

    temp["Clase de movimiento"]="101"
    del temp["Status de sistema"]
    del temp["Status"]
    del temp["Fecha liberac.real"]

    temp["Valor de la toma (WAERS)"]=0.0
    temp["Precio/MonL (WAERS)"]=0.0

    dfMb51["Clase"]="Cabecera"

    dfMb51=dfMb51.rename(columns={"Importe ML":"Valor de la toma (WAERS)",
                          "Ctd.en UM entrada":"Cantidad tomada (EINHEIT)",
                          "Unidad medida base":"Unidad medida base (=EINHEIT)"})

    dfMb51["Precio/MonL (WAERS)"]=dfMb51["Valor de la toma (WAERS)"].divide(dfMb51["Cantidad tomada (EINHEIT)"],fill_value=0)
    dfMb51["Precio/MonL (WAERS)"].replace([np.inf, -np.inf], 0, inplace=True)

    dfMb51["Cantidad necesaria (EINHEIT)"]=dfMb51["Cantidad tomada (EINHEIT)"]
    dfMb51=dfMb51.merge(temp[["Material","Texto breve material"]].drop_duplicates(),on="Material",how="left")
    
    dfComp=pd.concat([dfComp,dfMb51],sort=False)
    del temp
    del dfMb51

    if not dfComp.shape[0]==dfComp.merge(dfCab,how="left",on=['Centro', 'Orden']).shape[0]:
        raise Exception("Filas añadidas")

    dfComp=dfComp.merge(dfCab,how="left",on=['Centro', 'Orden'])

    dfComp["Cantidad Escala"]=dfComp["Cantidad Real PT"].divide(dfComp["Cantidad Plan PT"],fill_value=0)*dfComp["Cantidad necesaria (EINHEIT)"]
    dfComp["Cantidad Escala"].replace([np.inf, -np.inf], 0, inplace=True)

    dfComp=dfComp.rename(columns={"Precio/MonL (WAERS)":"Precio Plan"})

    #Precios en 0

    dfM=pd.read_excel(rutaM+"\MM60.xlsx",usecols=["Material","Centro","Cantidad base","Precio","Control de precios"],
                      converters={"Material":str,"Centro":str})
    dfM=dfM.drop_duplicates(subset=["Material","Centro"])

    dfM=dfM.rename(columns={"Precio":"Precio MM60"})
    dfM["Precio MM60"]=dfM["Precio MM60"].divide(dfM["Cantidad base"],fill_value=0).fillna(0).replace([np.inf, -np.inf], 0)

    if not dfComp.shape[0]==dfComp.merge(dfM,on=["Centro","Material"],how="left").shape[0]:
        raise Exception("MM60 inserta datos")

    dfComp=dfComp.merge(dfM,on=["Centro","Material"],how="left")

    dfComp["Cantidad base"].fillna(1,inplace=True)
    dfComp["Precio Plan"]=dfComp["Precio Plan"].divide(dfComp["Cantidad base"]).replace([np.inf, -np.inf], 0)

    dfComp["Precio tomado"]=dfComp["Valor de la toma (WAERS)"].divide(dfComp["Cantidad tomada (EINHEIT)"],fill_value=0).replace([np.inf, -np.inf], 0)
    dfComp["Precio tomado"].fillna(0,inplace=True)

    dfComp["Precio esperado"] = dfComp["Precio Plan"]
    dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio esperado"]=dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio tomado"]
    
    dfMb51=pd.read_excel(ruta+"\{}\Consumos\{}. MB51 (Consumos).xlsx".format(tiempo[0],tiempo[1]),usecols=colsMb51,converters=convMb51)
    dfMb51["Clase de movimiento"]=dfMb51["Clase de movimiento"].replace(["262","102","532"],["261","101","531"])
    
    dfMb51=dfMb51.groupby(["Orden","Centro","Material","Clase de movimiento"]).sum().reset_index()
    dfMb51["Precio MB51"]=dfMb51["Importe ML"].divide(dfMb51["Ctd.en UM entrada"],fill_value=0).fillna(0).replace([np.inf, -np.inf], 0)
    if not dfComp.shape[0]==dfComp.merge(dfMb51[["Orden","Centro","Material","Clase de movimiento","Precio MB51"]],
                                  on=["Orden","Centro","Material","Clase de movimiento"],
                                 how="left").shape[0]:
        raise Exception("Mb51 añade valores")
    dfComp=dfComp.merge(dfMb51[["Orden","Centro","Material","Clase de movimiento","Precio MB51"]],
                                  on=["Orden","Centro","Material","Clase de movimiento"],
                                 how="left")

    dfComp["Precio MB51"].fillna(0,inplace=True)
    dfComp["Precio MM60"].fillna(0,inplace=True)

    dfComp.loc[dfComp["Precio esperado"]==0.0,"Precio esperado"]=dfComp.loc[dfComp["Precio esperado"]==0.0,"Precio MB51"]
    dfComp.loc[dfComp["Precio esperado"]==0.0,"Precio esperado"]=dfComp.loc[dfComp["Precio esperado"]==0.0,"Precio MM60"]
    dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio Plan"]=dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio MB51"]
    dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio Plan"]=dfComp.loc[dfComp["Precio Plan"]==0.0,"Precio MM60"]
    dfComp.loc[dfComp["Precio tomado"]==0.0,"Precio tomado"]=dfComp.loc[dfComp["Precio tomado"]==0.0,"Precio MB51"]    
    dfComp.loc[dfComp["Precio tomado"]==0.0,"Precio tomado"]=dfComp.loc[dfComp["Precio tomado"]==0.0,"Precio MM60"]

    dfComp["Costo Tomado"]=dfComp["Cantidad tomada (EINHEIT)"]*dfComp["Precio tomado"]
    dfComp.loc[dfComp["Costo Tomado"]==0.0,"Costo Tomado"]=dfComp.loc[dfComp["Costo Tomado"]==0.0,"Valor de la toma (WAERS)"]
    dfComp.loc[dfComp["Valor de la toma (WAERS)"]==0.0,"Valor de la toma (WAERS)"]=dfComp.loc[dfComp["Valor de la toma (WAERS)"]==0.0,"Costo Tomado"]

    dfComp["Costo Estándar"]=dfComp["Cantidad Escala"]*dfComp["Precio Plan"]
    dfComp["Costo Esperado"]=dfComp["Cantidad tomada (EINHEIT)"]*dfComp["Precio esperado"]

    dfComp["Variación Consumo"]=dfComp["Costo Esperado"]-dfComp["Costo Estándar"]
    dfComp["Variación Precio"]=dfComp["Costo Tomado"]-dfComp["Costo Esperado"]
    del dfComp['Cantidad base']

    dfComp=dfComp[['Orden', 'Centro', 'Receta', 'Desc. Receta',"Clase",'Clase de movimiento', 
                       'Material', 'Texto breve material','Unidad medida base (=EINHEIT)',
                       'Cantidad necesaria (EINHEIT)','Cantidad tomada (EINHEIT)',
                        'Valor de la toma (WAERS)', 'Fecha liberac.real', 'Status',
                       'Cantidad Plan PT', 'Cantidad Real PT', 'Cantidad Escala',
                       'Precio tomado', 'Precio esperado', 'Precio Plan',
                        'Costo Estándar', 'Costo Esperado',"Costo Tomado",
                   'Variación Consumo', 'Variación Precio',"Control de precios"]]

    dfEj=pd.read_excel(ruta+"\{}\Ejecución\{}. Cuenta 7 Industria.xlsx".format(tiempo[0],tiempo[1]),usecols=colsEjec,converters=convEjec)

    dfEj=dfEj[dfEj["Clase de coste"].isin(["PP1001","PP1002"])]

    dfEj.loc[:,"Valor de la toma (WAERS)"]=-dfEj["Valor/mon.inf."]
    dfEj.loc[:,"Variación Precio"]=-dfEj["Desv.precio fija"]
    dfEj.loc[:,"Cantidad necesaria (EINHEIT)"]=-dfEj["Cantidad total"]

    del dfEj["Valor/mon.inf."]
    del dfEj["Desv.precio fija"]
    del dfEj["Cantidad total"]

    dfEj=dfEj.groupby(["Clase de coste","Denom.clase de coste","Orden partner"]).sum().reset_index()

    dfEj=dfEj.rename(columns={"Clase de coste":"Material","Denom.clase de coste":"Texto breve material","Orden partner":"Orden"})

    if not dfEj.shape[0]==dfEj.merge(dfCab,on=["Orden"],how="left").shape[0]:
        raise Exception("Cab inserta filas")

    dfEj=dfEj.merge(dfCab,on=["Orden"],how="left")

    dfEj["Unidad medida base (=EINHEIT)"]="H"

    dfEj["Cantidad tomada (EINHEIT)"]=dfEj["Cantidad necesaria (EINHEIT)"]
    dfEj["Cantidad Plan MP"]=dfEj["Cantidad necesaria (EINHEIT)"]

    dfEj['Precio tomado']=dfEj['Valor de la toma (WAERS)'].divide(dfEj['Cantidad necesaria (EINHEIT)'],fill_value=0).fillna(0)

    for i in ['Precio esperado', 'Precio Plan']:
        dfEj[i]=dfEj['Precio tomado']

    dfEj['Costo Estándar']=dfEj['Valor de la toma (WAERS)']-dfEj['Variación Precio']
    dfEj['Costo Esperado']=dfEj['Costo Estándar']
    dfEj['Costo Tomado']=dfEj['Valor de la toma (WAERS)']
    dfEj['Variación Consumo']=0

    for i in dfEj.columns:
        if i not in dfComp.columns:
            del dfEj[i]

    dfEj["Clase"]="Conversión"
    dfEj["Clase de movimiento"]="CONV"

    dfEj["Var Cuenta 7"]=dfEj["Variación Precio"]
    dfEj["Factor"]=dfEj['Cantidad Real PT'].divide(dfEj['Cantidad Plan PT'],fill_value=0).fillna(0)
    dfEj["Variación Consumo"]=(dfEj["Valor de la toma (WAERS)"]-dfEj["Variación Precio"])*(1-dfEj["Factor"])
    dfEj["Cantidad Escala"]=dfEj["Factor"]*dfEj["Cantidad tomada (EINHEIT)"]
    del dfEj["Factor"]

    dfComp["Var Cuenta 7"]=0

    dfComp=pd.concat([dfComp,dfEj])
    
    
    colsTemp=["Cantidad necesaria (EINHEIT)","Cantidad tomada (EINHEIT)","Valor de la toma (WAERS)","Cantidad Escala",
             "Costo Estándar","Costo Esperado","Costo Tomado"]
    dfComp.loc[dfComp["Clase"]=="Cabecera",colsTemp]=dfComp.loc[dfComp["Clase"]=="Cabecera",colsTemp]*-1
    
    
    vT=dfComp[["Orden","Costo Tomado","Variación Consumo","Variación Precio","Receta"]].copy()
    vT=vT.groupby(["Orden","Receta"]).sum().reset_index()

    vT=vT.rename(columns={"Costo Tomado":"Variación Total"})

    vT["Otras Variaciones"]=vT["Variación Total"]-vT["Variación Consumo"]-vT["Variación Precio"]

    vT["Clase"]="Cabecera"

    if not dfComp.shape[0]==dfComp.merge(vT[["Orden","Receta","Clase","Variación Total","Otras Variaciones"]],how="left",on=["Orden","Receta","Clase"]).shape[0]:
        raise Exception("vT inserta filas")

    dfComp=dfComp.merge(vT[["Orden","Receta","Clase","Variación Total","Otras Variaciones"]],how="left",on=["Orden","Receta","Clase"])
    dfComp["Variación Total"].fillna(0,inplace=True)
    dfComp["Otras Variaciones"].fillna(0,inplace=True)

    dfComp=dfComp.rename(columns={"Variación Total":"Neto Orden"})

    dfCentro=pd.read_excel(rutaM+"\Centros.xlsx",usecols=["Centro","Descripción Centro"],converters={"Centro":str})

    dfComp=dfComp.merge(dfCentro,how="left",on=["Centro"])


    dfM=dfM[["Material","Centro","Control de precios"]].rename(columns={"Material":"Receta",
                                                                        "Control de precios":"Control de precios Receta"})

    if not dfComp.shape[0]==dfComp.merge(dfM,on=["Receta","Centro"],how="left").shape[0]:
        raise Exeption("MM60 inserta filas")

    dfComp=dfComp.merge(dfM,on=["Receta","Centro"],how="left")

    for i in dfComp.columns:
        if dfComp[i].dtype == "object":
            dfComp[i].fillna("No encontrado",inplace=True)
        if (dfComp[i].dtype == "float64") or (dfComp[i].dtype == "int64"):
            dfComp[i].fillna(0.0,inplace=True)
        if (dfComp[i].dtype == "datetime64[ns]") or (dfComp[i].dtype == "int64"):
            dfComp[i].fillna(datetime(tiempo[0],tiempo[1],1),inplace=True)

    dfComp["Variación Total"]=dfComp["Variación Consumo"]+dfComp['Variación Precio']+dfComp['Otras Variaciones']

    cols=['Orden', 'Centro', 'Descripción Centro', 'Receta', 'Desc. Receta', 'Clase', 'Clase de movimiento',
       'Material', 'Texto breve material', 'Unidad medida base (=EINHEIT)',
       'Cantidad necesaria (EINHEIT)', 'Cantidad tomada (EINHEIT)',
       'Valor de la toma (WAERS)', 'Fecha liberac.real',
       'Status', 'Cantidad Plan PT', 'Cantidad Real PT', 'Cantidad Escala',
       'Precio tomado', 'Precio esperado', 'Precio Plan', 
          'Costo Estándar', 'Costo Esperado', 'Costo Tomado',
       'Variación Consumo', 'Variación Precio',  'Otras Variaciones',"Variación Total",'Neto Orden',"Var Cuenta 7",
         'Control de precios','Control de precios Receta']

    dfComp=dfComp[cols]
    dfComp=dfComp[~((dfComp["Receta"]==dfComp["Material"]) & (dfComp["Clase de movimiento"]=="101")& (dfComp["Clase"]=="Otras Entregas"))]
    dfComp.to_excel(rutaR+"\Consumos\{}\{}. Consumos.xlsx".format(tiempo[0],tiempo[1]),index=None)

    print(tiempo)

    print("{} {} Consumos generados con éxito".format(y,m))




def closeExcel(fname,qExcel=parse_args().qExcel):
    
    time.sleep(10)
    xl=win32com.client.Dispatch("Excel.Application")
    #print("Libros: {}".format(len(xl.Workbooks)))
    for wb in xl.Workbooks:
        #print(wb.Name)
        if wb.Name ==fname:
            wb.Close()
            wb=None
    if qExcel:
        xl.Quit()
    xl=None
    



def getReport(args,ruta=ruta):
    
    if args.consumo or args.ejec or args.ejecCEBE or args.maestra or args.mb51 or args.mb51B or args.prod or args.despachos or args.ke24:
        session=sapConnection(False)
    
    if args.consumo:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            cooisCabeceras(session,m,y,ruta)
            cooisComponentes(session,m,y,ruta)
            cooisAdicionales(session,m,y,ruta)
            consumosMB51(session,m,y,ruta)
            ksb1(session,m,y,ruta)
            produccion(session,m,y,ruta)
            produccionCarnes(session,m,y,ruta)
            maestras(session,ruta)
            reporteConsumos(m,y,ruta)
            correos.correoC7()
            correos.correoConsumos()
    
    if args.consumoR:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            reporteConsumos(m,y,ruta)
            
    if args.ejec:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            ksb1(session,m,y,ruta)
            produccion(session,m,y,ruta)
            produccionCarnes(session,m,y,ruta)
            
    if args.ejecCEBE:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            cebeC7(session,m,y,ruta)
            cebeCV(session,m,y,ruta)
            cebeIng(session,m,y,ruta)
            cebeGasto(session,m,y,ruta)
    
    if args.maestra:
        maestras(session,ruta)
    
    if args.mb51:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            consumosMB51(session,m,y,ruta)
        
    if args.mb51B:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            bajasMB51(session,m,y,ruta)
    
    if args.prod:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            produccion(session,m,y,ruta)
            produccionCarnes(session,m,y,ruta)

    if args.despachos:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            despachoKg(session,m,y,ruta)

    if args.ke24:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            ke24(session,m,y,ruta)
        
if __name__ == "__main__":
    
    args=parse_args()
    getReport(args)