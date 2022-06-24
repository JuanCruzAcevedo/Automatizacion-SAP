import os
import win32com.client as win32
import pandas as pd
from Herramientas_normalizadoras import Herramientas_normalizadoras

class Sap():
    def __init__(self,credenciales = ""):
        SapGui = win32.GetObject("SAPGUI").GetScriptingEngine
        self.session = SapGui.FindById("ses[0]")
        if credenciales == "":
            self.credenciales = credenciales
        else:
            self.credenciales = credenciales
            self.herramientas = Herramientas_normalizadoras(self.credenciales)
    
    def guardar(self,ruta):
        """Guarda los archivos, se le indica el nombre del archivo y la ruta a guardar""" 
        self.session.findById("wnd[0]").sendVKey(16)
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            os.remove("{}".format(ruta))
            print("borrado el archivo viejo")
        except:
            print('...')
                    
        try:
            print('Se esta guardando el archivo')
            xl = win32.dynamic.Dispatch('Excel.Application')
            wb = xl.Workbooks('Hoja de cálculo en Basis (1)').SaveAs(Filename="{}".format(ruta))
            xl.Workbooks('Hoja de cálculo en Basis (1)').Close()
            print('Se  guardo el archivo')
        except:
            print('No se encontro ni guardo el archivo')
    
    def transaccion(self,transaccion):
        '''Abre una transaccion'''
        self.session.StartTransaction(Transaction=transaccion)
        
    def ingresar_variante(self,variante,f8=False):
        '''Ingresa y busca una variante, en caso de poner True en el 
        parametro F8, '''
        self.session.findById("wnd[0]").sendVKey(17)
        self.session.findById("wnd[1]/usr/txtV-LOW").text = variante
        self.session.findById("wnd[1]/usr/txtV-LOW").caretPosition =(13)
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]").sendVKey(8)
        if f8 == False:
            pass
        elif f8 == True:
            self.session.findById("wnd[0]").sendVKey(8)
        else:
            raise Exception('Ingrese un valor valido')
        
    def copiar_avisos(self,ruta,columna):
        """Copia los avisos de un excel y los pone en un porta papeles"""
        avisos = pd.read_excel("{}".format(ruta))
        avisos = avisos[[columna]]
        avisos.to_clipboard(index=0)
    
    def buscar_avisos(self,transaccion,ruta_avisos):
        self.copiar_avisos(ruta_avisos,'Aviso')
        self.transaccion(transaccion)
        self.ingresar_variante('/AVISOS_561')
        self.session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]").sendVKey(24)
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)
    
    def descargar_datos_ciudadanos(self,ruta_copiar,ruta_guardar,nombre):
        '''En el parametro 'ruta_copiar' se pone la ruta de donde sacar los avisos, en "ruta_guardar" se coloca el solamente la ruta donde se guardara,
        y  por ultimo en "nombre" se seleciona el nombre que tendra el archivo '''
        self.transaction("ZDATOS_CIUDADANOS")
        self.copiar_avisos(ruta_copiar,'Aviso')
        self.session.findById("wnd[0]/usr/btn%_SO_QMNUM_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/tbar[1]/btn[46]").press()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "%pc"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = "{}".format(ruta_guardar)
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}.txt".format(nombre)
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
        df1=pd.read_table("{}{}.txt".format(ruta_guardar,nombre),sep="\t",encoding="latin-1", error_bad_lines=False)
        df1['Nombre del Ciudadano'] = df1['Apellido']+', '+df1['Nombre']
        df1.rename(columns = {'Cor.elec.':'Mail','Tel.Cel':'Tel celular','Teléfono':'Tel fijo'},inplace=True)
        df1.rename(columns = {'Dir.cor.elec.':'Mail','Teléfono Celular':'Tel celular','Teléfono':'Tel fijo'},inplace=True)
        df1.to_excel("{}{}.xlsx".format(ruta_guardar,nombre), index = False)
        os.remove("{}{}.txt".format(ruta_guardar,nombre))

    def descarga_descripcion(self,avisos_buscar,ruta_guarda,nombre):
        """Copia los avisos del archivo que se le indique en 'avisos buscar', luego se le indica 
        la ruta a guardar de los archivos en 'ruta guarda' (solo la ruta) y finalemente se ingresa el nombre para el archivo"""
        self.transaccion("ZR010_TAVISOS")
        self.copiar_avisos(avisos_buscar,'Aviso')
        self.session.findById("wnd[0]/usr/radRB_MAN").select()
        self.session.findById("wnd[0]/usr/btn%_P_AVISOS_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/tbar[1]/btn[46]").press()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "%pc"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = "{}".format(ruta_guarda)
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}.txt".format(nombre)
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
        #----normaliza el txt-------#
        df1=pd.read_table("{}{}.txt".format(ruta_guarda,nombre),sep="\t",encoding="latin-1", error_bad_lines=False) 

        df1=df1[["Objeto","Texto Extendido"]]

        df1=df1.groupby('Objeto').agg(lambda x: x.tolist())

        df1['Texto Extendido']=df1['Texto Extendido'].astype(str)
        textos=["\['\* ","\)', '\*","', '\*","'\]","', '  ",","]

        for texto in textos:
            df1['Texto Extendido']=df1['Texto Extendido'].str.replace("{}".format(texto),' ')

        df1.to_excel("{}{}.xlsx".format(ruta_guarda,nombre))
        os.remove("{}{}.txt".format(ruta_guarda,nombre))

    def crear_pedido(self,mes,texto_breve,ruta, guardar = False):
        if self.credenciales == "":
            print("Es necesario crear la clase con credenciales")
            
        else:
            self.transaccion("ME21N")
            #pone pedido marco GCBA
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").setFocus()
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").key = "FO"

            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "100195" #provedor
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = "gcba"#org compras
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = "ev1"#grupo compras
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text = "gcba"#sociedad
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1229/ctxtMEPO1229-KDATB").text = self.herramientas.meses.get(mes).get('inicio extremo')
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1229/ctxtMEPO1229-KDATE").text = self.herramientas.meses.get(mes).get('fin extremo')
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]").text = "u"
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,0]").text = "d"
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]").text = texto_breve
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[14,0]").text = "SERV-MTTO"
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").text = "GCBA"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1").select()
            
            
            
            #Lee el archivo de Excel y corrige las Unidades de Medida
            u_medida_corregir = {
                'ud':'un','gl':'un',
                'jornal':'un','dia':'un'
            }
            
            df = pd.read_excel(ruta)
            df['Unidad de Medida'] = df['Unidad de Medida'].map(u_medida_corregir).fillna(df['Unidad de Medida'])
            
            #Delimita el final y el inicio por index dividiendolos de a 10 
            df.index = df.index + 1
            cantidad = int(df.index.max()/10)
            resto = df.index.max()%10
            inicio,fin = [],[]

            for i in range(cantidad):
                if i+1 == 1:
                    inicial= (i+1)
                    final = (i+1)*10
                    inicio.append(inicial)
                    fin.append(final)
                else:
                    inicial= (i)*10
                    final = (i+1)*10
                    inicio.append(inicial)
                    fin.append(final)

            inicio.append(cantidad*10)
            fin.append(cantidad*10+resto)
            
            #carga los valores filtrando los listados 
            for numero,valor in enumerate(inicio):
                if numero == 0:
                    carga = df.loc[(df.index<=10)]
                    for x,a in enumerate(carga['Clave']):
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-KTEXT1[3,{}]".format(x)).text =carga['Clave'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-MENGE[4,{}]".format(x)).text =carga['Duracion'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-MEINS[5,{}]".format(x)).text =carga['Unidad de Medida'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-TBTWR[6,{}]".format(x)).text =carga['Precio Neto'].tolist()[x]

                    #baja el cursor
                    self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW").verticalScrollbar.position = 10

                else:
                    carga = df.loc[(df.index>inicio[numero])&(df.index<=fin[numero])]

                    for x,a in enumerate(carga['Clave']):
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-KTEXT1[3,{}]".format(x+1)).text =carga['Clave'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-MENGE[4,{}]".format(x+1)).text =carga['Duracion'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-MEINS[5,{}]".format(x+1)).text =carga['Unidad de Medida'].tolist()[x]
                        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-TBTWR[6,{}]".format(x+1)).text =carga['Precio Neto'].tolist()[x]
                        
                    #baja el cursor
                    self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW").verticalScrollbar.position = fin[numero]
                            
            
            
            precio_neto = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").text
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT2").select()
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT2/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1327/subSUB0:SAPLMLSP:0401/subLIMIT:SAPLMLSL:0115/txtESUH-SUMLIMIT").text = precio_neto
            self.session.findById("wnd[0]").sendVKey(0)
            if guardar == False:
                print('Para que se guarde poner en el parametro "guardar" True')
                pass 
            elif guardar == True:
                self.session.findById("wnd[0]").sendVKey(11)
                self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                numero_pedido = self.session.findById("wnd[0]/sbar").text.split()[-1]
                print('el numero de pedido marco es :\n{}'.format(numero_pedido))
                
    def liquidacion(self,ruta):        
        #primera parte vincula los avisos
        self.transaccion('ZC011_VINC')
        self.session.findById("wnd[0]/usr/ctxtPA_PATH").text = ruta
        self.session.findById("wnd[0]").sendVKey(8)
        #clikea el warning en caso de que la orden se encuentre asociada
        while True:
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                break

        print('Asocio las ordenes')
        #verifica que los montos sean los correspondientes
        self.transaccion('Z_REP_OPERACIONES')
        pass
    #solamente llega a asociar las ordenes 
    
