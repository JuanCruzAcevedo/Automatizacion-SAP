import os
import win32com.client as win32
import pandas as pd

class Sap():
    def __init__(self):
        SapGui = win32.GetObject("SAPGUI").GetScriptingEngine
        self.session = SapGui.FindById("ses[0]")
    
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
    
