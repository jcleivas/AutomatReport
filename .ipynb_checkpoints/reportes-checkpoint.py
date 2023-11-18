import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def cmv(x):
    if x["Clase de movimiento"] in ["101","102"]:
        return "Otras Entregas"
    elif x["Clase de movimiento"] in ["261","262"]:
        return "Consumos"
    elif x["Clase de movimiento"] in ["531","532"]:
        return "Subproductos"
    else:
        return "No encontrado"

def cantidadEscalada(x):
    if x["Clase"]=="Cabecera":
        return x["Cantidad necesaria (EINHEIT)"]
    else:
        return x["Cantidad Real PT"]/x["Cantidad Plan PT"]*x["Cantidad necesaria (EINHEIT)"]
    
def reporteConsumos(m,y,ruta,rutaM,rutaR):
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

    dfComp["Cantidad Escala"]=dfComp.apply(cantidadEscalada,axis=1)
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

    print("{} {} Consumos generados con éxito".format(y,m))