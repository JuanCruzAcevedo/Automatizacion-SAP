from  Herramientas_normalizadoras import Herramientas_normalizadoras
from Cargas_Drive import Archivos_drive
import pandas as pd


class Normalizar_archivos():
    def __init__(self):
        self.herramientas = Herramientas_normalizadoras()
        pass

    def normalizar_r11(self):
        """Normaliza el listado R11 y devuelve varios DF """
        #Trae del drive el archivo
        df_relevamientos = Archivos_drive(r"C:\Users\20407295650\Desktop\credenciales.json","R11","Respuestas de formulario 1")#hacer variable el path
        relevamientos = df_relevamientos.abrir_archivo()
        #filtra el archivo a las columnas que precisamos y normaliza DAP y Altura
        filtrado = relevamientos[['Marca temporal', 'Nombre completo', 'Numero de aviso', 'Calle',
               'Chapa', 'Referencia', 'Situación de la posición', 'Especie',
               'DAP (cm)', 'Altura (m)','Tarea Recomendada','Tipo de poda','Tipo de corte de raíz', 'Condición del árbol',
               'Riesgo', 'Prioridad', 'Observaciones', 'Status', 'id',
               'Status Avisos ', 'Orden', 'Liquidacion']]
        filtrado['Altura (m)']=pd.to_numeric(filtrado['Altura (m)'],errors = 'coerce')
        filtrado['DAP (cm)'] =pd.to_numeric(filtrado['DAP (cm)'],errors = 'coerce')

        #Si no tiene hay numero de avisos lo filtra del listado 

        filtrado = filtrado[filtrado['Numero de aviso']!=""]
        filtrado['Numero de aviso'] = filtrado['Numero de aviso'].str.upper()

        #Crea listado de Avisos Oficio y Presidencia

        avisos_oficio_presidencia = filtrado.loc[(filtrado['Numero de aviso']=='OFICIO')|(filtrado['Numero de aviso']=='PRESIDENCIA')]

        #Crea listado de avisos duplicados y los ordena alfabeticamente

        avisos_duplicados = filtrado.drop(avisos_oficio_presidencia.index.to_list(), axis=0)
        avisos_duplicados = avisos_duplicados[avisos_duplicados.duplicated(subset = ['Numero de aviso'],keep=False)]
        avisos_duplicados = avisos_duplicados.sort_values('Numero de aviso')


        #Crea listado de avisos final sin incluir los listados anteriores ni tampoco los avisos sin id, Status ERROR y a denegar

        final = filtrado.drop((avisos_oficio_presidencia.index.to_list()+avisos_duplicados.index.to_list()), axis=0)
        final = final.loc[final['id']!='']
        final = final[final['Status Avisos ']!='---ERROR---']
        final = final[final['Status Avisos ']!='']
        final = final.loc[~final['Status Avisos '].isin(self.herramientas.sacar_status)]
        
        #Crea listado de avisos a denegar 
        avisos_denegar = final[final['Tarea Recomendada'].isin(['','No corresponde ninguna tarea'])]

        #Termina de crear el listado final
        final = final.drop(avisos_denegar.index.to_list(), axis=0)
    
        #A partir de este punto trabaja con el listado final, y le suma las Claves Modelo
        #Suma Clave modelo de Poda (usa apply)
        final.loc[filtrado['Tarea Recomendada']=='Poda','Clave Modelo']= final['Tipo de poda']\
            .apply(self.herramientas.podas).astype(str)+final['Altura (m)'].apply(self.herramientas.altura_podas).astype(str)
        final.loc[final['Clave Modelo'].str.len() == 9,'Clave Modelo'] = 'AR-P401'

        #Suma Clave Modelo Corte de Raiz(usa .map)

        final.loc[filtrado['Tarea Recomendada']=='Corte de  Raíces','Clave Modelo']=final['Tipo de corte de raíz'].map(self.herramientas.corte_raices).fillna('ERROR') 

        #Suma Clave Modelo de Plantacion

        final.loc[final['Tarea Recomendada']=='Plantar','Clave Modelo']='AR-PL03'

        #Suma 'Agrandar Plantera' en Clave Modelo

        final.loc[final['Tarea Recomendada']=='Agrandar plantera','Clave Modelo']='Agrandar Plantera'

        #Suma Clave Modelo de Extracciones (usa .map y .apply) 

        final.loc[final['Tarea Recomendada']=='Extracción','Clave Modelo']= final['Condición del árbol'].map(self.herramientas.extracciones).fillna('ERROR')+final['DAP (cm)'].apply(self.herramientas.dap)
        final.loc[final['Clave Modelo'].isin(['AR-RC07','AR-RC08']),'Clave Modelo']='AR-RC09'
        #Devuelve los DF
        
        warning ="EL SCRIPT NO ESTA TERMINADO, FALTA QUE A LAS PODAS LES SUMEN SUS RETIROS \n"
        print(warning,'Se devolveran los DF en el siguiente orden:\n"avisos oficio/presidencia"\n"Duplicados"\n"Para Denegar"\n"Final para cargar"')
        
        return [avisos_oficio_presidencia,avisos_duplicados,avisos_denegar,final]
    
    def normalizar_mt1(self,df):
        pass