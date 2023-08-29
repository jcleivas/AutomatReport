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
import numpy as np
import warnings
import win32com.client
import os


def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 80:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

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
    parser.add_argument("-mdr",dest="mdr",action="store_true",help="Genera el reporte Modelo de Rentabilidad")
    parser.add_argument("-f",dest="fechas",nargs=4,help="Permite indicar de qué período es el reporte, si no se incluye se descargaran los reportes del mes en curso")
    dia=datetime.now()
    parser.set_defaults(pdl=False,desp=False,ing=False,mdr=False,fechas=[dia.month,dia.year,dia.month+1,dia.year])
    args=parser.parse_args()
    return args


def MDRReport(m,y):
    rCv=r"C:\Users\jcleiva\Documents\Reportes Base\{}\Costo de Ventas\{}. Costo de Ventas Industria CEBE (Agg).xlsx"
    rCvT=r"C:\Users\jcleiva\Documents\Reportes Base\{}\Costo de Ventas\{}. Costo de Ventas Industria CEBE.xlsx"
    rPLUIND=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\P&G Industria\V3\Insumo Plantas Industria Despacho-Venta.xlsx"
    rMM60=r"C:\Users\jcleiva\Documents\Reportes Base\Maestras\MM60.xlsx"
    rIng=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\Ingresos\{}{:02d}. Ingresos.xlsx"
    rLPYG=r"C:\Users\jcleiva\Documents\Reportes Base\Maestras\Línea P&G.xlsx"
    rCEBE=r"C:\Users\jcleiva\Documents\Reportes Base\Maestras\Maestro CEBE Industria.xlsx"
    rGasto=r"C:\Users\jcleiva\Documents\Reportes Base\{}\Gasto\{}. Gasto Industria CEBE.xlsx"
    cLog=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\P&G Industria\V3\Costo Logístico\Costo Logístico.xlsx"
    rCentros=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\Maestras\Centros.xlsx"
    rDesp=r"C:\Users\jcleiva\Documents\Reportes Base\{}\Despachos\{}. Despachos {}Q.xlsx"
    rMDR=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\MDR\{}{:02d}. MDR v4.xlsx"
    rPDL=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\PDL\{}{:02d}. PDL.xlsx"    

    vCols=['Fecha','Centro de beneficio', 'Número de cuenta', 'Denominación Cuenta',
       'Material', 'Centro', 'Importe', 'Cantidad', 
       'PLU Industria', 'PLU Venta', 'Sublinea', 'Desc. Sublinea',
       'Desc. Plu Venta', 'Unitario Ingreso (PLU Venta)',"PDL",
       'Unitario Costo (PLU Venta)', 'Desc. Plu Industria',
       'Unitario Ingreso (PLU Industria)', 'Unitario Costo (PLU Industria)',
       'Costo Producto', 'Venta Bruta','Venta Neta', 'Denominación CEBE',
       'Bajas / Averías_(Dir)', 'Depreciación_(Dir)',
       'Devoluciones Almacenes_(Dir)', 'Merma_(Dir)', 'Variación_(Dir)',
       'Bajas / Averías_(Mat)', 'Depreciación_(Mat)',
       'Devoluciones Almacenes_(Mat)', 'Merma_(Mat)', 'Variación_(Mat)',
       'Bajas / Averías_(Sub)', 'Depreciación_(Sub)',
       'Devoluciones Almacenes_(Sub)', 'Merma_(Sub)', 'Variación_(Sub)',
       'Bajas / Averías_(Tr)', 'Depreciación_(Tr)',
       'Devoluciones Almacenes_(Tr)', 'Merma_(Tr)', 'Variación_(Tr)',
       'Depreciación Gasto Industria', 'Gasto Industria',
       'Costo Logístico_(Dir)', 'Costo Logístico_(Mat)',
       'Costo Logístico_(Sub)', 'Costo Logístico_(Tr)',"Texto breve de material","Descripción Centro",
      'Destinatario de mercancías','Desc Destinatario de mercancías', 'Marca Formato']

    colsCebe=["Centro de beneficio","Número de cuenta","Denominación",
              "Importe","Cantidad","Material","Centro"]
    convCebe={"Centro de beneficio":str,"Número de cuenta":str,"Material":str,"Centro":str}
    colsPYG=["Número de Cuenta","Línea P&G"]
    convPYG={"Número de Cuenta":str}
    colsIng=["Sublinea","Desc. Sublinea","Plu","Desc. Plu","# Unidades Totales","$ Ventas sin impuestos","$ Costo"]
    convIng={"Sublinea":str,"Plu":str}

    colsDesp=["Centro de beneficio","Número de cuenta","Denominación","Material","Centro","En moneda local centro de beneficio","Cantidad","Documento comercial"]
    convDesp={"Centro de beneficio":str,"Número de cuenta":str,"Material":str,"Centro":str,"Documento comercial":str}

    cDesp=["Destinatario de mercancías","Entrega","    # PLU","Material","Marca Formato"]
    cvDesp={"Entrega":str,"    # PLU":str,"Material":str}

    colsPDL=["Plu","Dependencia","$ Precio Venta Historico"]
    cvPDL={"Plu":str,"Dependencia":str}
    
    #Costo de Ventas
    dfCV=pd.read_excel(rCv.format(y,m),converters=convCebe,usecols=colsCebe)
    dfCV=dfCV.rename(columns={"Denominación":"Denominación Cuenta"})
    dfCV=dfCV[~dfCV["Centro de beneficio"].isin(["7670","7680"])]
    dfCV["Material"].fillna("No encontrado",inplace=True)
    dfCV["Centro"].fillna("No encontrado",inplace=True)
    dfM=pd.read_excel(rLPYG,usecols=colsPYG,converters=convPYG)
    dfM=dfM.rename(columns={"Número de Cuenta":"Número de cuenta"})
    dfCV=dfCV.merge(dfM,on="Número de cuenta",how="left")
    dfCV.loc[dfCV["Número de cuenta"].isin(["612014012","612014014"]),"Línea P&G"]="Devoluciones Almacenes"
    dfCV=dfCV[~dfCV["Línea P&G"].isna()]
    dfDesp=pd.read_excel(rCvT.format(y,m,1),converters=convDesp,usecols=colsDesp)
    dfDesp=dfDesp[dfDesp["Número de cuenta"]=="612014014"]
    dfDesp=dfDesp[~dfDesp["Centro de beneficio"].isin(["7670","7680"])]
    dfDesp["Línea P&G"]='Costo Despacho'
    dfDesp=dfDesp.rename(columns={"En moneda local centro de beneficio":"Importe","Denominación":"Denominación Cuenta"})
    dfDesp=dfDesp.groupby(['Centro de beneficio', 'Número de cuenta', "Denominación Cuenta",
                           'Material', 'Centro','Documento comercial', 'Línea P&G']).sum().reset_index()
    dfDespQ=pd.read_excel(rDesp.format(y,m,1),converters=cvDesp,usecols=cDesp)
    dfDespQ=pd.concat([dfDespQ,pd.read_excel(rDesp.format(y,m,2),converters=cvDesp,usecols=cDesp)])
    dfDespQ=dfDespQ.drop_duplicates()
    dfDespQ=dfDespQ.rename(columns={"Entrega":"Documento comercial","    # PLU":"PLU Industria"})
    if not dfDesp.merge(dfDespQ,on=["Documento comercial","Material"],how="left").shape[0]==dfDesp.shape[0]:
        raise Exception("Despachos por Q inserta filas")
    dfDesp=dfDesp.merge(dfDespQ,on=["Documento comercial","Material"],how="left")

    del dfDesp["Documento comercial"]
    dfDesp["PLU Industria"]=dfDesp["PLU Industria"].str.lstrip("0")

    dfM=pd.read_excel(rPLUIND,usecols=["Plu Despacho","Plu Venta"],
                      converters={"Plu Despacho":str,"Plu Venta":str},sheet_name="Insumo Venta-Despacho")

    dfM=dfM.drop_duplicates(subset="Plu Despacho")
    dfDesp=dfDesp.merge(dfM.rename(columns={"Plu Despacho":"PLU Industria"}),on="PLU Industria",how="left")

    dfDesp["Plu Venta"].fillna(dfDesp["PLU Industria"],inplace=True)
    dfDesp=dfDesp.rename(columns={"Plu Venta":"PLU Venta"})
    
    #Ingreso

    dfIng=pd.read_excel(rIng.format(y,y,m),usecols=colsIng, converters=convIng)

    dfIng=dfIng.rename(columns={"Plu":"PLU Venta","# Unidades Totales":"Cantidad","$ Ventas sin impuestos":"Ingreso","$ Costo":"Costo"})

    dfIng=dfIng.groupby(["Sublinea","Desc. Sublinea","PLU Venta","Desc. Plu"],dropna=False).sum().reset_index()

    dfIng["Unitario Ingreso"]=dfIng["Ingreso"].divide(dfIng["Cantidad"],fill_value=0.0)
    dfIng["Unitario Ingreso"].replace([np.inf, -np.inf], 0, inplace=True)
    dfIng["Unitario Ingreso"].fillna(0.0,inplace=True)

    dfIng["Unitario Costo"]=dfIng["Costo"].divide(dfIng["Cantidad"],fill_value=0.0)
    dfIng["Unitario Costo"].replace([np.inf, -np.inf], 0, inplace=True)
    dfIng["Unitario Costo"].fillna(0.0,inplace=True)

    del dfIng["Ingreso"]
    del dfIng["Costo"]
    del dfIng["Cantidad"]

    if dfDesp.merge(dfIng,on="PLU Venta",how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas")

    dfDesp=dfDesp.merge(dfIng.rename(columns={'Desc. Plu':'Desc. Plu Venta',
                                              "Unitario Ingreso":"Unitario Ingreso (PLU Venta)",
                                             "Unitario Costo":"Unitario Costo (PLU Venta)"}),on="PLU Venta",how="left")

    if dfDesp.merge(dfIng.rename(columns={"PLU Venta":"PLU Industria"}),on="PLU Industria",how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas")

    del dfIng['Sublinea']
    del dfIng['Desc. Sublinea']

    dfDesp=dfDesp.merge(dfIng.rename(columns={"PLU Venta":"PLU Industria",'Desc. Plu':'Desc. Plu Industria',
                                             "Unitario Ingreso":"Unitario Ingreso (PLU Industria)",
                                             "Unitario Costo":"Unitario Costo (PLU Industria)"}),on="PLU Industria",how="left")

    for i in dfDesp.columns:
        if dfDesp[i].dtype == "object":
            dfDesp[i].fillna("No encontrado",inplace=True)
        if dfDesp[i].dtype == "float64":
            dfDesp[i].fillna(0,inplace=True)

    dfDesp["Costo Producto"]=dfDesp.apply(costoIndustria,axis=1)
    dfDesp["Venta Neta"]=dfDesp["Unitario Ingreso (PLU Venta)"]*dfDesp["Cantidad"]

    dfDesp["Costo Producto"]=dfDesp.apply(costoIndustria,axis=1)
    dfM=pd.read_excel(rCEBE,usecols=["Centro de beneficio","Denominación CEBE"],converters={"Centro de beneficio":str})
    dfM=dfM.drop_duplicates(subset=["Centro de beneficio"])
    dfDesp=dfDesp.merge(dfM,on="Centro de beneficio",how="left")

    dfVN=dfDesp[["Centro de beneficio","Denominación CEBE","Material","Venta Neta"]].groupby(["Centro de beneficio","Denominación CEBE","Material"]).sum().reset_index().copy()

    dfCV=dfCV.merge(dfM,on="Centro de beneficio",how="left")
    
    #Sacrificio a carnes

    dfCV["Centro de beneficio"].replace("7850","7852",inplace=True)
    dfCV["Denominación CEBE"].replace("Planta Sacrificio Re","Planta Carnes Centro",inplace=True)

    dfCV=dfCV.groupby(["Centro de beneficio","Número de cuenta","Denominación Cuenta","Material",
                 "Centro","Línea P&G","Denominación CEBE"],dropna=False).sum().reset_index()

    #Otras Líneas Costo

    dfCV=dfCV[~dfCV["Línea P&G"].isin(["Costo Despacho","Costo Venta a Terceros"])]

    del dfCV["Denominación CEBE"]
    del dfCV["Centro"]

    dfCV=pd.pivot_table(dfCV,values=["Importe"], index=["Centro de beneficio","Material"],
                       columns=["Línea P&G"], aggfunc='sum',dropna=False).reset_index().fillna(0)

    cTemp=[]
    for i in dfCV.columns:
        if i[1]!="":
            cTemp.append(i[1])
        else:
            cTemp.append(i[0])


    dfCV.columns=cTemp

    dfCV=dfCV[~((dfCV["Bajas / Averías"]==0.0)&(dfCV["Merma"]==0.0)&(dfCV["Variación"]==0.0)&(dfCV["Depreciación"]==0.0))]

    del dfVN["Denominación CEBE"]

    dfCV=dfCV.merge(dfVN[["Centro de beneficio","Material","Venta Neta"]].groupby(["Centro de beneficio","Material"]).sum(),
                    how="left",on=["Centro de beneficio","Material"])

    temp=dfCV[(~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0)]
    dfCV=dfCV[~((~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0))]

    for i in temp.columns:
        if i not in ["Centro de beneficio","Material","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["Centro de beneficio","Material"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 2")
    dfDesp=dfDesp.merge(temp,on=["Centro de beneficio","Material"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Dir)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    # material
    del dfCV["Venta Neta"]
    dfCV=dfCV.merge(dfVN[["Material","Venta Neta"]].groupby(["Material"]).sum(),on="Material",how="left")

    temp=dfCV[(~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0)]
    dfCV=dfCV[~((~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0))]

    del temp["Centro de beneficio"]
    del temp["Venta Neta"]

    temp=temp.groupby(["Material"]).sum().reset_index()

    temp=temp.merge(dfVN[["Material","Venta Neta"]].groupby(["Material"]).sum(),on="Material",how="left")

    for i in temp.columns:
        if i not in ["Centro de beneficio","Material","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["Material"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 3")
    dfDesp=dfDesp.merge(temp,on=["Material"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Mat)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    del dfCV["Venta Neta"]

    # Centro de beneficio
    del dfCV["Material"]

    dfCV=dfCV.groupby(["Centro de beneficio"]).sum().reset_index()
    dfCV=dfCV.merge(dfVN[['Centro de beneficio','Venta Neta']].groupby(["Centro de beneficio"]).sum().reset_index(),on="Centro de beneficio",how="left")

    temp=dfCV[(~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0)]
    dfCV=dfCV[~((~dfCV["Venta Neta"].isna())&(dfCV["Venta Neta"]!=0.0))]

    for i in temp.columns:
        if i not in ["Centro de beneficio","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["Centro de beneficio"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 4")
    dfDesp=dfDesp.merge(temp,on=["Centro de beneficio"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Sub)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    del dfCV["Venta Neta"]

    dfCV["Venta Neta"]=dfVN["Venta Neta"].sum().copy()

    for i in dfCV.columns:
        if i not in ["Centro de beneficio","Venta Neta"]:
            dfCV[i+"_ratio"]=dfCV[i].divide(dfCV["Venta Neta"])
            del dfCV[i]
            dfDesp[i+"_(Tr)"]=dfCV[i+"_ratio"].sum()*dfDesp["Venta Neta"]

    dfDesp.fillna(0,inplace=True)

    #Gasto
    dfGasto=pd.read_excel(rGasto.format(y,m),
                          usecols=["Número de cuenta","Centro de beneficio","En moneda local centro de beneficio"],
                          converters={"Número de cuenta":str})

    dfGasto["Tipo"]=dfGasto["Número de cuenta"].apply(lambda x: "Depreciación Gasto Industria" if x[:4] in ["5160","5260","5265"] else "Gasto Industria")
    #dfGasto.loc[dfGasto["Número de cuenta"].isin(["529525003"]),"Tipo"]="Capex"
    dfGasto=dfGasto[~dfGasto["Número de cuenta"].isin(["529525003"])] #Capex
    dfGasto=dfGasto[~dfGasto["Centro de beneficio"].isin(["7670","7680"])]
    del dfGasto["Centro de beneficio"]
    dfGasto=dfGasto[["Tipo","En moneda local centro de beneficio"]].groupby("Tipo").sum().reset_index()

    for i in dfGasto["Tipo"]:
        ratio=dfGasto.loc[dfGasto["Tipo"]==i,"En moneda local centro de beneficio"].sum() / dfVN["Venta Neta"].sum().copy()
        dfDesp[i]=ratio*dfDesp["Venta Neta"]

    dfCLog=pd.read_excel(cLog,usecols=["Año","Mes","Sublinea","Plu","Costo Logístico"],
                        converters={"Sublinea":str,"Plu":str})

    dfCLog=dfCLog[dfCLog["Año"]==y]
    dfCLog=dfCLog[dfCLog["Mes"]==m]
    del dfCLog["Mes"]
    del dfCLog["Año"]

    dfCLog=dfCLog.groupby(["Sublinea","Plu"]).sum().reset_index()

    dfCLog=dfCLog.rename(columns={"Plu":"PLU Venta"})

    dfCLog=dfCLog[~((dfCLog["Costo Logístico"]==0.0))]
    dfVN=dfDesp[["Sublinea","PLU Venta","Venta Neta"]].groupby(["Sublinea","PLU Venta"]).sum().reset_index().copy()

    dfCLog=dfCLog.merge(dfVN,how="left",on=["Sublinea","PLU Venta"])

    temp=dfCLog[(~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0)]
    dfCLog=dfCLog[~((~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0))]

    for i in temp.columns:
        if i not in ["Sublinea","PLU Venta","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["Sublinea","PLU Venta"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 2")
    dfDesp=dfDesp.merge(temp,on=["Sublinea","PLU Venta"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Dir)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    # material
    del dfCLog["Venta Neta"]
    dfCLog=dfCLog.merge(dfVN[["PLU Venta","Venta Neta"]].groupby(["PLU Venta"]).sum(),on="PLU Venta",how="left")

    temp=dfCLog[(~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0)]
    dfCLog=dfCLog[~((~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0))]

    del temp["Sublinea"]
    del temp["Venta Neta"]

    temp=temp.groupby(["PLU Venta"]).sum().reset_index()
    temp=temp.merge(dfVN[["PLU Venta","Venta Neta"]].groupby(["PLU Venta"]).sum(),on="PLU Venta",how="left")

    for i in temp.columns:
        if i not in ["Sublinea","PLU Venta","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["PLU Venta"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 3")
    dfDesp=dfDesp.merge(temp,on=["PLU Venta"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Mat)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    del dfCLog["Venta Neta"]

    # Sublinea
    del dfCLog["PLU Venta"]

    dfCLog=dfCLog.groupby(["Sublinea"]).sum().reset_index()

    dfCLog=dfCLog.merge(dfVN[['Sublinea','Venta Neta']].groupby(["Sublinea"]).sum().reset_index(),on="Sublinea",how="left")

    temp=dfCLog[(~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0)]
    dfCLog=dfCLog[~((~dfCLog["Venta Neta"].isna())&(dfCLog["Venta Neta"]!=0.0))]

    for i in temp.columns:
        if i not in ["Sublinea","Venta Neta"]:
            temp[i+"_ratio"]=temp[i].divide(temp["Venta Neta"])
            del temp[i]

    del temp["Venta Neta"]

    if dfDesp.merge(temp,on=["Sublinea"],how="left").shape[0]!=dfDesp.shape[0]:
        raise Exception("Inserta filas 4")
    dfDesp=dfDesp.merge(temp,on=["Sublinea"],how="left")

    dfDesp.fillna(0,inplace=True)
    for i in dfDesp.columns:
        if "_ratio" in i:
            dfDesp[i[:-6]+"_(Sub)"]=dfDesp[i]*dfDesp["Venta Neta"]
            del dfDesp[i]

    del dfCLog["Venta Neta"]
    if dfCLog.shape[0]>0:
        dfCLog["Venta Neta"]=dfCLog["Venta Neta"].sum().copy()

        for i in dfCLog.columns:
            if i not in ["Sublinea","Venta Neta"]:
                dfCLog[i+"_ratio"]=dfCLog[i].divide(dfCLog["Venta Neta"])
                del dfCLog[i]
                dfDesp[i+"_(Tr)"]=dfCLog[i+"_ratio"].sum()*dfDesp["Venta Neta"]

        dfDesp.fillna(0,inplace=True)
    else:
        dfDesp["Costo Logístico_(Tr)"]=0

    dfDesp["Fecha"]=datetime.datetime(y,m,1)
    dfM=pd.read_excel(rMM60,usecols=["Material","Texto breve de material"],converters={"Material":str})
    dfM=dfM.drop_duplicates(subset=["Material"])
    dfDesp=dfDesp.merge(dfM,on=["Material"],how="left")
    dfM=pd.read_excel(rCentros,usecols=["Centro","Descripción Centro"],converters={"Centro":str})
    dfM=dfM.drop_duplicates(subset=["Centro"])
    dfDesp=dfDesp.merge(dfM,on=["Centro"],how="left")

    del dfDesp["Línea P&G"]

    dfDesp[["Destinatario de mercancías","Desc Destinatario de mercancías"]]=dfDesp["Destinatario de mercancías"].str.split("-",n=1,expand=True)
    dfDesp["Destinatario de mercancías"]=dfDesp["Destinatario de mercancías"].str.strip()
    dfDesp["Desc Destinatario de mercancías"]=dfDesp["Desc Destinatario de mercancías"].str.strip()


    dfPDL=pd.read_excel(rPDL.format(y,y,m),converters=cvPDL,usecols=colsPDL)

    dfPDL=dfPDL.rename(columns={"Plu":"PLU Venta","Dependencia":"Destinatario de mercancías","$ Precio Venta Historico":"PDL"})

    if not dfDesp.merge(dfPDL,how="left",on=["PLU Venta","Destinatario de mercancías"]).shape[0]==dfDesp.shape[0]:
        raise Exception("PDL inserta filas")

    dfDesp=dfDesp.merge(dfPDL,how="left",on=["PLU Venta","Destinatario de mercancías"])
    dfDesp["PDL"].fillna(dfDesp["Unitario Ingreso (PLU Venta)"],inplace=True)

    dfDesp["Venta Bruta"]=dfDesp["PDL"]*dfDesp["Cantidad"]


    if not set(dfDesp.columns) == set(vCols):
        raise Exception("Existen columnas diferentes a las permitidas")
        print(list(set(dfDesp.columns) - set(vCols)))


    defCols=['Fecha','Centro de beneficio','Denominación CEBE', 'Número de cuenta', 'Denominación Cuenta',
       'Material', 'Texto breve de material', 'Centro','Descripción Centro', 'Importe', 'Cantidad', 
       'PLU Industria', 'Desc. Plu Industria','Unitario Ingreso (PLU Industria)', 'Unitario Costo (PLU Industria)',
      'PLU Venta', 'Desc. Plu Venta', 'Unitario Ingreso (PLU Venta)','Unitario Costo (PLU Venta)',"PDL",
      'Sublinea', 'Desc. Sublinea','Destinatario de mercancías',"Desc Destinatario de mercancías", 'Marca Formato',
        'Venta Bruta','Venta Neta', 'Costo Producto',  
       'Bajas / Averías_(Dir)', 'Depreciación_(Dir)',
       'Devoluciones Almacenes_(Dir)', 'Merma_(Dir)', 'Variación_(Dir)',
       'Bajas / Averías_(Mat)', 'Depreciación_(Mat)',
       'Devoluciones Almacenes_(Mat)', 'Merma_(Mat)', 'Variación_(Mat)',
       'Bajas / Averías_(Sub)', 'Depreciación_(Sub)',
       'Devoluciones Almacenes_(Sub)', 'Merma_(Sub)', 'Variación_(Sub)',
       'Bajas / Averías_(Tr)', 'Depreciación_(Tr)',
       'Devoluciones Almacenes_(Tr)', 'Merma_(Tr)', 'Variación_(Tr)',
       'Depreciación Gasto Industria', 'Gasto Industria',
       'Costo Logístico_(Dir)', 'Costo Logístico_(Mat)',
       'Costo Logístico_(Sub)', 'Costo Logístico_(Tr)']

    for i in dfDesp.columns:
        if i in ['Venta Bruta',"Importe",'Venta Neta', 'Costo Producto', 'Bajas / Averías_(Dir)', 'Depreciación_(Dir)',
                   'Devoluciones Almacenes_(Dir)', 'Merma_(Dir)', 'Variación_(Dir)',
                   'Bajas / Averías_(Mat)', 'Depreciación_(Mat)','Devoluciones Almacenes_(Mat)', 'Merma_(Mat)', 'Variación_(Mat)',
                   'Bajas / Averías_(Sub)', 'Depreciación_(Sub)','Devoluciones Almacenes_(Sub)', 'Merma_(Sub)', 'Variación_(Sub)',
                   'Bajas / Averías_(Tr)', 'Depreciación_(Tr)', 'Devoluciones Almacenes_(Tr)', 'Merma_(Tr)', 'Variación_(Tr)',
                   'Depreciación Gasto Industria', 'Gasto Industria','Costo Logístico_(Dir)', 'Costo Logístico_(Mat)',
                   'Costo Logístico_(Sub)', 'Costo Logístico_(Tr)']:
            dfDesp[i]=dfDesp[i]/1000000


    dfDesp=dfDesp[defCols]
    dfDesp["Desc. Plu Industria"]=dfDesp["Desc. Plu Industria"].str.capitalize()
    dfDesp["Desc. Plu Venta"]=dfDesp["Desc. Plu Venta"].str.capitalize()
    dfDesp.to_excel(rMDR.format(y,y,m),index=None)
    print("Reporte MDR {}{} generado con éxito".format(m,y))
    
    
    
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
        columns=["Plu","Desc. Plu",'Dependencia',"Desc. Dependencia","Proveedor Plu-Dep",
                 "Desc. Proveedor Plu-Dep","Dia",
                 "Estado PluDepHistoria","Desc. Estado PluDepHistoria","$ Precio Venta Historico",
                 "$ CPM Historico","$ Precio Fabrica Historico","$ Costo Neto Historico","$ Costo Sugerido Historico"]
        conv={"Plu":str,"Dependencia":str,"Proveedor Plu-Dep":str}
        rows=9
        
    if tipo=="Ingresos":
        columns=["Sublinea","Desc. Sublinea","Clase Marca","Marca","Desc. Marca","Plu","Desc. Plu","Formato",
                 "Desc. Formato","Cadena","Desc. Cadena","Dependencia","Desc. Dependencia","Proveedor","Desc. Proveedor",
                 "# Unidades Totales","$ Ventas sin impuestos","$ Costo"]
        rows=6
        conv={"Sublinea":str,"Marca":str,"Plu":str,"Dependencia":str,"Proveedor":str}
        
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        pd.read_excel(join(path,filename),skiprows=rows,header=None,names=columns,converters=conv).to_excel(join(path,filename),index=None)
    print("Titulos corregidos {}".format(filename))
        
def pdlReport(driver,m,y,mypathD,mypathPDL):
    
    element = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr247"]/div[2]/div/div/div[2]/div/a[4]')))
    element.click()
    #driver.find_element(By.XPATH,'//*[@id="id_mstr247"]/div[2]/div/div/div[2]/div/a[4]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr264_txt"]').clear()
    driver.find_element(By.XPATH,'//*[@id="id_mstr264_txt"]').send_keys("01/{:02d}/{}".format(m,y))
    element = WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr266"]')))
    element.click()
    #driver.find_element(By.XPATH,'//*[@id="id_mstr266"]').click()

    driver.find_element(By.XPATH,'//*[@id="id_mstr253"]').click()

    original_window=driver.current_window_handle
    ###
    element = WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tbExport"]')))
    element.click()
    
    driver.switch_to.window(driver.window_handles[-1])

    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]
    
    element = WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="3131"]')))
    element.click()
    
    element = WebDriverWait(driver, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="done_eb_ExportStyle"]/a')))
    #element.click()
    
    download_wait(mypathD)
    time.sleep(20)
    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. PDL.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    print("Reporte de pdl descargado: {} {}".format(y,m))
    
    titlesReport(mypathPDL, "{}{:02d}. PDL.xlsx".format(y,m),"PDL")

def despReport(driver,m,y,mypathD,mypathPDL):

    driver.find_element(By.XPATH,'//*[@id="id_mstr207"]/div[2]/div/div/div[2]/div/a[4]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr303_txt"]').clear()
    meses={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
          11:"Noviembre",12:"Diciembre"}
    driver.find_element(By.XPATH,'//*[@id="id_mstr303_txt"]').send_keys("{} {}".format(meses[m],y))
    
    driver.find_element(By.XPATH,'//*[@id="id_mstr305"]').click()
    driver.find_element(By.XPATH,'//*[@id="id_mstr292"]').click()
    original_window=driver.current_window_handle
    
    element = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tbExport"]')))
    element.click()
    driver.switch_to.window(driver.window_handles[-1])

    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]
    
    
    element = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportPageByInfo"]')))
    element.click()
    driver.find_element(By.XPATH,'//*[@id="exportReportTitle"]').click()
    driver.find_element(By.XPATH,'//*[@id="exportFilterDetails"]').click()
    driver.find_element(By.XPATH,'//*[@id="3131"]').click()
    
    element = WebDriverWait(driver, 80).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="done_eb_ExportStyle"]/div[1]')))
    element.click()
    download_wait(mypathD)

    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. Despachos.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    print("Reporte de Despachos descargado: {} {}".format(y,m))
    titlesReport(mypathPDL, "{}{:02d}. Despachos.xlsx".format(y,m),"Despachos")

def ingReport(driver,m,y,mypathD,mypathPDL):
    driver.find_element(By.XPATH,'//*[@id="id_mstr116"]/table/tbody/tr[8]/td[2]').click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr329"]/div[2]/div/div/div[2]/div/a[4]')))
    element.click()
    meses={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
          11:"Noviembre",12:"Diciembre"}
    
    try:
        driver.find_element(By.XPATH,'//*[@id="id_mstr370_txt"]').clear() 
        driver.find_element(By.XPATH,'//*[@id="id_mstr370_txt"]').send_keys("{} {}".format(meses[m],y))
        driver.find_element(By.XPATH,'//*[@id="id_mstr372"]').click()
        print(370)
    except:
        try:
            driver.find_element(By.XPATH,'//*[@id="id_mstr460_txt"]').clear()
            driver.find_element(By.XPATH,'//*[@id="id_mstr460_txt"]').send_keys("{} {}".format(meses[m],y))
            driver.find_element(By.XPATH,'//*[@id="id_mstr462"]').click()        
            print(460)
        except:
            try:
                driver.find_element(By.XPATH,'//*[@id="id_mstr415_txt"]').clear()
                driver.find_element(By.XPATH,'//*[@id="id_mstr415_txt"]').send_keys("{} {}".format(meses[m],y))
                driver.find_element(By.XPATH,'//*[@id="id_mstr417"]').click()
                print(415)
            except:
                try:
                    driver.find_element(By.XPATH,'//*[@id="id_mstr340_txt"]').clear()
                    driver.find_element(By.XPATH,'//*[@id="id_mstr340_txt"]').send_keys("{} {}".format(meses[m],y))
                    driver.find_element(By.XPATH,'//*[@id="id_mstr342"]').click()
                    print(340)
                except:
                    input("Hola")
    
    driver.find_element(By.XPATH,'//*[@id="id_mstr284"]').click()

    original_window=driver.current_window_handle

    element = WebDriverWait(driver, 80).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tbExport"]')))
    element.click()

    driver.switch_to.window(driver.window_handles[-1])
    prefiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]
    
    element = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="3131"]')))
    element.click()
    
    element = WebDriverWait(driver, 80).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="done_eb_ExportStyle"]/div[1]')))
    element.click()
    download_wait(mypathD)
    
    posfiles = [f for f in listdir(mypathD) if isfile(join(mypathD, f)) and f[-5:]==".xlsx"]

    nFiles=list(set(posfiles) - set(prefiles))
    if len(nFiles)==1:
        shutil.copy(join(mypathD, nFiles[0]),join(mypathPDL, "{}{:02d}. Ingresos.xlsx".format(y,m)))
    else:
        print("No pudimos determinar el archivo nuevo: {} {:02d}".format(y,m))

    driver.close()

    driver.switch_to.window(original_window)
    driver.find_element(By.XPATH,'//*[@id="tbBack0"]').click()

    print("Reporte de Ingresos descargado: {} {}".format(y,m))
    titlesReport(mypathPDL, "{}{:02d}. Ingresos.xlsx".format(y,m),"Ingresos")
    
def getReportSinemax(args):
    
    if args.pdl:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\PDL"
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
            pdlReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL.format(tiempo[0]))
        
    if args.desp:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\Despachos"
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
            despReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL.format(tiempo[0]))
    
    if args.ing:
        path=r"C:\Users\jcleiva\OneDrive - Grupo-exito.com\Escritorio\Proyectos\AutomatizacionExito\chromedriver_win32\chromedriver.exe"
        mypathD=r"C:\Users\jcleiva\Downloads"
        mypathPDL=r"C:\Users\jcleiva\Documents\Reportes Base\{}\P&G\Ingresos"
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
            ingReport(driver,tiempo[1],tiempo[0],mypathD,mypathPDL.format(tiempo[0]))
    
    if args.mdr:
        for tiempo in month_year_iter(int(args.fechas[0]),int(args.fechas[1]),int(args.fechas[2]),int(args.fechas[3])):
            MDRReport(tiempo[1],tiempo[0])
    
if __name__ == "__main__":
    args=parse_args()
    getReportSinemax(args)