import subprocess
import shutil
import time
import win32com.client
import pythoncom
import argparse
from calendar import monthrange
from datetime import datetime, timedelta

def month_year_iter( start_month, start_year, end_month, end_year ):
    ym_start= 12*start_year + start_month - 1
    ym_end= 12*end_year + end_month - 1
    for ym in range( ym_start, ym_end ):
        y, m = divmod( ym, 12 )
        yield y, m+1

def parse_args():
    parser = argparse.ArgumentParser(description="Esta función determina qué reportes descargar")
    parser.add_argument("-c",dest="consumo",action="store_true",help="Descarga reportes de Consumo (MB51, COOISPI)")
    parser.add_argument("-e",dest="ejec",action="store_true",help="Descarga las ejecuciones")
    parser.add_argument("-ck",dest="costoKilo",action="store_true",help="Descarga Costo Kilo desde ZCO_COSTO_KILO")
    parser.add_argument("-cUmb",dest="costoUMB",action="store_true",help="Descarga Costo Kilo desde SQVI")
    parser.add_argument("-m",dest="maestra",action="store_true",help="Descarga Maestras (MM60, COPA, Versiones de Fabricación)")
    parser.add_argument("-f",dest="fechas",nargs=4,help="Permite indicar de qué período es el reporte, si no se incluye se descargaran los reportes del mes en curso")
    parser.add_argument("-ke",dest="ke30",action="store_true",help="Descarga Ke30")
    parser.add_argument("-keE",dest="ke30E",action="store_true",help="Descarga Ke30 Ecuador")
    parser.add_argument("-ka",dest="kardex",action="store_true",help="Descarga Kardex")
    parser.add_argument("-x",dest="copiar",action="store_true",help="Copia los reportes descargados en la ruta de Reportes Base")
    parser.add_argument("-qE",dest="qExcel",action="store_true",help="Cierra la aplicación Excel")
    dia=datetime.now()
    parser.set_defaults(consumo=False, ejec=False, costoKilo=False,maestra=False,costoUMB=False,fechas=[dia.month,dia.year,dia.month+1,dia.year],ke30E=False, ke30=False,copiar=False)
    args=parser.parse_args()
    return args
    

def sapConnection(cSap):
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
            
            connection = application.OpenConnection("ALPINA - Sistema S/4HANA PROD")
            time.sleep(5)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()
            return session            
            
        except:
            connection = application.OpenConnection("ALPINA - Sistema S/4HANA PROD")
            time.sleep(5)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()
            return session

    except pythoncom.com_error as error:
        hr,msg,exc,arg = error.args
        
        if "The 'Sapgui Component' could not be instantiated." == exc[2]:
            print("No se pudo iniciar SAP, revisa tu conexión a internet/VPN")
        else:
            print(exc[2])
            raise Exception(error)     

            
def bajarKe30(session,ruta):
        
    session.findById("wnd[0]/usr/radLISTE").select()
    session.findById("wnd[0]/usr/radLISTE").setFocus()

    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/lbl[1,9]").setFocus()
    session.findById("wnd[0]/usr/lbl[1,9]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/lbl[1,7]").setFocus()
    session.findById("wnd[0]/usr/lbl[1,7]").caretPosition = 11
    session.findById("wnd[0]").sendVKey(2)

    session.findById("wnd[0]/tbar[1]/btn[48]").press()
    session.findById("wnd[1]/usr/btnD2000_PUSH_01").press()
    session.findById("wnd[1]/tbar[0]/btn[6]").press()
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[0,0]").selected = True
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[7,0]").selected = True
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[2,0]").selected = True
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[1,0]").selected = True
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    xl=win32com.client.Dispatch("Excel.Application")
    act=xl.ActiveWorkbook
    act.Sheets(1).Name="Ke30Fam"
    wb=xl.Workbooks.Add()
    act.Sheets("Ke30Fam").Copy(wb.Sheets(1))

    for i in wb.Sheets:
        if i.Name != "Ke30Fam":
            xl.DisplayAlerts=False
            wb.Sheets(i.Name).Delete()
            xl.DisplayAlerts=True
    wb.SaveAs(Filename=ruta)
    wb.Close()

    session.findById("wnd[1]/tbar[0]/btn[0]").press()

def bajarKe30E(session,ruta):
        
    session.findById("wnd[0]/usr/radLISTE").select()
    session.findById("wnd[0]/usr/radLISTE").setFocus()

    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/lbl[1,9]").setFocus()
    session.findById("wnd[0]/usr/lbl[1,9]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/lbl[1,7]").setFocus()
    session.findById("wnd[0]/usr/lbl[1,7]").caretPosition = 11
    session.findById("wnd[0]").sendVKey(2)

    session.findById("wnd[0]/tbar[1]/btn[48]").press()
    session.findById("wnd[1]/usr/btnD2000_PUSH_01").press()
    session.findById("wnd[1]/tbar[0]/btn[6]").press()
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[5,0]").selected = True
    session.findById("wnd[1]/usr/sub:SAPLKEC1:0100/chkCEC01-CHOICE[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    xl=win32com.client.Dispatch("Excel.Application")
    act=xl.ActiveWorkbook
    act.Sheets(1).Name="Ke30Fam"
    wb=xl.Workbooks.Add()
    act.Sheets("Ke30Fam").Copy(wb.Sheets(1))

    for i in wb.Sheets:
        if i.Name != "Ke30Fam":
            xl.DisplayAlerts=False
            wb.Sheets(i.Name).Delete()
            xl.DisplayAlerts=True
    wb.SaveAs(Filename=ruta)
    wb.Close()

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
def ke30(session,m,y,ruta):
    k=0
    while k<10:
        r="{:012d}".format(1010+k)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKE30"
        session.findById("wnd[0]").sendVKey(0)

        try:
            session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "GALP"
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/shellcont/shell").doubleClickNode (r)
        except:
            print("No galp")
            session.findById("wnd[0]/shellcont/shell").doubleClickNode (r)
        
        try:
            session.findById("wnd[0]/usr/ctxtPAR_01").text = "{:03d}.{:04d}".format(m,y)
            session.findById("wnd[0]/usr/radLISTE").setFocus()

            print(session.ActiveWindow.Text)
            print("Reporte {}".format(r))
            if session.ActiveWindow.Text=="Selección: P&G UN COL":    
                k=11
                bajarKe30(session,ruta)
            else:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except:
            print("Reporte {} no existe".format(r))
        k=k+1
        print(k)

        
def ke30E(session,m,y,ruta):
    k=0
    while k<10:
        r="{:012d}".format(1010+k)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKE30"
        session.findById("wnd[0]").sendVKey(0)

        try:
            session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "GALP"
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/shellcont/shell").doubleClickNode (r)
        except:
            print("No galp")
            session.findById("wnd[0]/shellcont/shell").doubleClickNode (r)
        
        try:
            session.findById("wnd[0]/usr/ctxtPAR_01").text = "{:03d}.{:04d}".format(m,y)
            session.findById("wnd[0]/usr/radLISTE").setFocus()

            print(session.ActiveWindow.Text)
            print("Reporte {}".format(r))
            if session.ActiveWindow.Text=="Selección: PYG Ecuador":    
                k=11
                bajarKe30E(session,ruta)
            else:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except:
            print("Reporte {} no existe".format(r))
        k=k+1
        print(k)

        
def kardex(session,m,y,ruta):
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZCO_KARDEX"
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/usr/ctxtP_BUKRS").text = "4800"
    session.findById("wnd[0]/usr/txtPJAHRPER").text = "{:03d}.{:04d}".format(m,y)
    
    session.findById("wnd[0]/usr/btn%_S_SAKNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1405050000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1405100000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1435150000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1435200000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "1460050000"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "300024"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "300025"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "301225"

    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(4,"MCOD1")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "4"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta.format(y)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}. Kardex.xlsx".format(m)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    
    
def cooispiCabeceras(session,m,y,ruta,ruta2,xcopy): #m stands for month, y for year
    i=m
    j=y
    dias=monthrange(j,i)[1]
    diasMesAnt=(datetime(j,i,1)-timedelta(days=1)).day
    fechaIni=datetime(j,i,1)-timedelta(days=1)-timedelta(days=diasMesAnt)+timedelta(days=1)
    mesI=fechaIni.month
    yearI=fechaIni.year
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncooispi"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-LOW").text = "{:02d}.{:02d}.{}".format(1,mesI,yearI)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = 7
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. Cabeceras de orden LM.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    p1=ruta+"\{}\Consumos\{}".format(j,fname)
    p2=ruta2+"\Reportes Base {}\{}".format(j,fname)
    if xcopy:
        shutil.copyfile(p1,p2)
    
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))
    
def consumosMB51(session,m,y,ruta,ruta2,xcopy):
    i=m
    j=y
    dias=monthrange(j,i)[1]
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "CO1030611534"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(5,"BTEXT")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. MB51 (Consumos).xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    p1=ruta+"\{}\Consumos\{}".format(j,fname)
    p2=ruta2+"\Reportes Base {}\{}".format(j,fname)
    if xcopy:
        shutil.copyfile(p1,p2)    
    
    closeExcel(fname)    
    print("{} descargado con éxito".format(fname))

def cooispiNotif(session,m,y,ruta,ruta2,xcopy):
    j=y
    i=m
    dias=monthrange(j,i)[1]

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncooispi"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOR000"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-LOW").text = "{:02d}.{:02d}.{}".format(1,i,j)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTFR-HIGH").text = "{:02d}.{:02d}.{}".format(dias,i,j)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell (1,"MEINH")
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Consumos".format(j)
    fname="{}. Rechazos.xlsx".format(i)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    p1=ruta+"\{}\Consumos\{}".format(j,fname)
    p2=ruta2+"\Reportes Base {}\{}".format(j,fname)
    if xcopy:
        shutil.copyfile(p1,p2)    
    
    closeExcel(fname)
    print("{} descargado con éxito".format(fname))

def ksb1(session,m,y,ruta,ruta2,xcopy):
    j=y
    i=m
    dias=monthrange(j,i)[1]
    grupos={"COLECT_GEO":"ColecGeo","MANUFACT":"Manufact","COLECT_COL":"ColecCol","7708":"SopOp Ecu","6608":"SopOp Vzl","4809":"LogCol","4808":"SopOpCol","3009":"LogZf"}
    for g in grupos.keys():
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "GALP"
            session.findById("wnd[1]").sendVKey(0)
        except:
            session.findById("wnd[0]/usr/ctxtP_KOKRS").text = "GALP"
        
        session.findById("wnd[0]/usr/ctxtKSTGR").text = g
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
        
        p1=ruta+"\{}\Ejecución\{}".format(j,fname)
        p2=ruta2+"\Reportes Base {}\Ejecución Presupuestal\{}".format(j,fname)
        if xcopy:
            shutil.copyfile(p1,p2)    
        
        closeExcel(fname)
        print("{} descargado con éxito".format(fname))
    
def zcoCostoKilo(session,m,y,ruta):
    soc={"4800":["4801","4802","4803","4804","4806","4811","4812","4813","4814","4815","4816","4817","4821","4822","4823",
            "T001","T002","T003","T004","T005","T006","T007","T008","T009","T010","T011","T012","T013","T014","T015",
            "T016","T017","T018","T019","T020","T021","T022","T023","T024","483*","4842","4891"],"3000":["3006"],
    "7700":["7701","7702","7711"],"6600":["6601","6609","6611","6612","6613","6614"]}
    
    for s in soc.keys():
        for c in soc[s]:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZCO_COSTO_KILO"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = s
            session.findById("wnd[0]/usr/txtPA_BDATJ").text = str(y)
            session.findById("wnd[0]/usr/txtPA_POPER").text = str(m)
            session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").text = "*"
            session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = c
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            try: 
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(13,"MAKTX")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "13"
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Costo Kilo".format(y)
                if c=="483*":
                    c="483"
                fname="{}. Costo Kilo {}.XLSX".format(m,c)
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                
                closeExcel(fname)
                print("{} descargado con éxito".format(fname))
                
            except pythoncom.com_error as error:
                hr,msg,exc,arg = error.args
                if "The control could not be found by id." != exc[2]:
                    print(error)

def costoUMB(session,m,y,ruta,ruta2,xcopy):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nsqvi"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "ctostanumb"
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 10
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "*"
    session.findById("wnd[0]/usr/txtSP$00003-LOW").text = str(y)
    session.findById("wnd[0]/usr/txtSP$00004-LOW").text = str(m)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell (1,"PEINH")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\{}\Costo Kilo".format(y)
    fname="{}. Costo UMB.xlsx".format(m)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]").sendVKey(11)
    
    p1=ruta+"\{}\Costo Kilo\{}".format(y,fname)
    p2=ruta2+"\Reportes Base {}\Costo Kilo\{}".format(y,fname)
    if xcopy:
        shutil.copyfile(p1,p2)    
    
    closeExcel(fname)
    print("{} descargado con éxito".format(fname))
    
def maestras(session,ruta,ruta2,xcopy):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm60"
    session.findById("wnd[0]").sendVKey(0)
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
    
    p1=ruta+"\Maestras\{}".format(fname)
    p2=ruta2+"\Reportes Base\{}".format(fname)
    if xcopy:
        shutil.copyfile(p1,p2)
    
    closeExcel(fname)
    print("{} descargado con éxito".format(fname))
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSQ01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn("DBGBNUM")
    session.findById("wnd[1]/tbar[0]/btn[29]").press()
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "ZCO"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "ZCO00000000006"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "*"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(4,"NTGEW")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta + "\Maestras"
    fname="Maestra COPA.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    p1=ruta+"\Maestras\{}".format(fname)
    p2=ruta2+"\Reportes Base\{}".format(fname)
    if xcopy:
        shutil.copyfile(p1,p2)
    
    closeExcel(fname)
    print("{} descargado con éxito".format(fname))
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSQVI"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "VERFABDF"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "*"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(1,"STKTX")
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta+"\Maestras"
    fname="Versión Fabricación.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    p1=ruta+"\Maestras\{}".format(fname)
    p2=ruta2+"\Reportes Base\{}".format(fname)
    if xcopy:
        shutil.copyfile(p1,p2)
    
    closeExcel(fname)
    print("{} descargado con éxito".format(fname))
    
def closeExcel(fname,qExcel=parse_args().qExcel):
    time.sleep(5)
    xl=win32com.client.Dispatch("Excel.Application")
    for wb in xl.Workbooks:
        if wb.Name ==fname:
            wb.Close()
            wb=None
    if qExcel:
        xl.Quit()
    xl=None

def getReport(args,ruta,ruta2):
    session=sapConnection(False)
    if args.consumo:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            cooispiCabeceras(session,m,y,ruta,ruta2,args.copiar)
            consumosMB51(session,m,y,ruta,ruta2,args.copiar)
            cooispiNotif(session,m,y,ruta,ruta2,args.copiar)
            
    if args.ejec:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            ksb1(session,m,y,ruta,ruta2,args.copiar)
    
    if args.costoKilo:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            zcoCostoKilo(session,m,y,ruta)
    
    if args.costoUMB:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            costoUMB(session,m,y,ruta,ruta2,args.copiar)
    
    if args.maestra:
        maestras(session,ruta,ruta2,args.copiar)
    
    if args.ke30:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            fname=ruta+"\\{}\\KE30\\{} {}. ke30 Col.xlsx".format(y,y,m)
            ke30(session,m,y,fname)
            
    if args.ke30E:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            fname=ruta+"\\{}\\KE30\\{} {}. ke30 Ecu.xlsx".format(y,y,m)
            ke30E(session,m,y,fname)
    
    if args.kardex:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            m=tiempo[1]
            y=tiempo[0]
            fname=r"C:\Users\juan.leiva\Desktop\Temp\TEMP\{}\Kardex"
            kardex(session,m,y,fname)
    
    session=sapConnection(True)
    
if __name__ == "__main__":
    ruta=r"C:\Users\juan.leiva\Desktop\Temp\TEMP"
    ruta2=r"C:\Users\juan.leiva\Desktop\Alpina\Proyectos"
    args=parse_args()
    getReport(args,ruta,ruta2)