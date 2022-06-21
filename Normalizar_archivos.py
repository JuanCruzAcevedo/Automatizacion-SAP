from  Herramientas_normalizadoras import Herramientas_normalizadoras
from Cargas_Drive import Archivos_drive
import pandas as pd
import numpy as np

class Normalizar_archivos():
    def __init__(self,credenciales):
        self.herramientas = Herramientas_normalizadoras(credenciales)
        self.credenciales = credenciales

    def normalizar_r11(self):
        """Normaliza el listado R11 y devuelve varios DF """
        #Trae del drive el archivo
        df_relevamientos = Archivos_drive(self.credenciales,"R11","Respuestas de formulario 1")#hacer variable el path
        relevamientos = df_relevamientos.abrir_archivo()
        #filtra el archivo a las columnas que precisamos y normaliza DAP y Altura
        filtrado = relevamientos[['Marca temporal', 'Nombre completo', 'Numero de aviso', 'Calle',
               'Chapa', 'Referencia', 'Situación de la posición', 'Especie',
               'DAP (cm)', 'Altura (m)','Tarea Recomendada','Tipo de poda','Tipo de corte de raíz', 'Condición del árbol',
               'Riesgo', 'Prioridad', 'Observaciones', 'Status', 'id',
               'Status Avisos ', 'Orden', 'Liquidacion']].copy()
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
    
    def normalizar_mt1(self,mes,df):
        """Se le pasa el mes en el que se pretende hacer la carga y el df correspondiente a MT1"""
        #resetea el index y lo toma como orden
        df.reset_index(inplace = True,drop = True)
        df['orden'] = df.index
        #suma las prestaciones y los campos que se resuelven facil
        df['Prestacion'] = df['Clave Modelo'].map(self.herramientas.prestaciones).fillna('NO ES UNA CLAVE MODELO')
        df['Cuadra (Automatico) '] = df['Chapa'].apply(self.herramientas.redondear)
        df['Corredor (automatico)'] = df['Calle']+" "+df['Cuadra (Automatico) ']
        df['orden1'] = 'ARRE'
        df['clase de actividad'] = 'MT1'
        df['grupo plani'] = 'AR1'
        df['Clave de campo'] = 'ARB-DG'
        df['Duracion'] = 1
        df['Fecha'] = '-'
        df['Especie'] = df['Especie'].str[:20].fillna('-')
        df['Especie'] = df['Especie'].replace('','-')
        
        
        #suma los campos inicio, fin y puesto de trabajo
        df['inicio extremo'] = self.herramientas.meses.get(mes).get('inicio extremo')
        df['fin extremo'] = self.herramientas.meses.get(mes).get('fin extremo')
        df['Emplazamiento'] = self.herramientas.meses.get(mes).get('campo clasi')
        
        #define las ubicaciones tecnicas
        df['Altura exacta'] = df['Calle']+" "+df['Chapa']
        ubt_exactas = self.herramientas.ubicacion_tecnica()
        ubt_corredores = self.herramientas.ubicacion_tecnica(False)

        df['Ubicación Tecnica'] = df['Altura exacta'].map(ubt_exactas).fillna(df['Corredor (automatico)'].map(ubt_corredores).fillna('No esta ubt'))

        #Carga el puesto de trabajo 

        df['puesto de trabajo'] = df['Nombre completo'].map(self.herramientas.inspectores).fillna('CUADC11')

        #Carga el texto breve

        df['texto breve'] = df['Calle']+' '+df['Chapa']+'-'+df['Prestacion'].map(self.herramientas.prestaciones_simp)

        #Les asigna nuevamente el nombre a las columnas

        df.rename(columns = {'Nombre completo':'Inspectores','Calle':'Calle real ','Referencia':'Ref','DAP (cm)':'DAP','Altura (m)':'Altura'}, inplace = True )
        
        #en DAP y Altura los valores nan los completa con "-"
        df.fillna(value ={'DAP':'-','Altura':'-'}, inplace = True,)
        
        #Crea una nueva columna para definir si es mas largo o no 
        df['Largo'] = df['texto breve'].str.len()>40
        df['Largo'] = np.where(df['Largo']==True,'Es un texto > 40','Es un texto < a 40')

        #Nos quedamos con las columnas que nos importan
        df = df[['orden', 'Fecha', 'Inspectores', 'Corredor (automatico)', 'Calle real ',
               'Cuadra (Automatico) ', 'orden1', 'Ubicación Tecnica', 'inicio extremo',
               'fin extremo', 'clase de actividad', 'puesto de trabajo', 'grupo plani',
               'texto breve', 'Clave de campo', 'Clave Modelo', 'Duracion', 'Chapa',
               'Ref', 'Especie', 'DAP', 'Altura', 'Emplazamiento','Prestacion','Largo']]
        

        #creamos las clave modelo de retiro para las podas
        retiro_de_poda = df.loc[df['Prestacion']=='Poda'].copy()

        condiciones = [
            (retiro_de_poda['Altura']<=12),
            (retiro_de_poda['Altura']>12) & (retiro_de_poda['Altura']<=20),
            (retiro_de_poda['Altura']>20)]

        valores = ['AR-RP01','AR-RP02','AR-RP03']

        retiro_de_poda['Clave Modelo'] = np.select(condiciones,valores)

        #creamos las clave modelo para las extracciones, corte de raiz y retiro de cepa
        veredas  = df.loc[df['Prestacion'].isin(['Corte de Raiz','Retiro de Cepa','Extraccion'])].copy()
        planteras = veredas.copy()

        veredas['Clave Modelo'] = 'AR-RV35'
        veredas['Duracion'] = 8

        planteras['Clave Modelo'] = 'AR-RV37'
        planteras['Duracion'] = 4

        #concatena los DF
        df = pd.concat([veredas,planteras,retiro_de_poda,df])

        #ordena el DF por numero de orden
        df.sort_values(by=['orden'],inplace = True)
        
       
        print('''
        {} lineas con textos que pasan los 40
        {} lineas sin la ubt 
        {} sin prestacion'''.format(
        df.loc[df['Largo']=='Es un texto > 40','Largo'].count(),
        df.loc[df['Ubicación Tecnica']=='No esta ubt','Ubicación Tecnica'].count(),
        df.loc[df['Prestacion']=='NO ES UNA CLAVE MODELO','Prestacion'].count()
        )
             )
        #hacer un print que diga cuantos hay sin ubt, cuantos sin prestacion y cuantos con el texto > a 40 
        return df

    def normalizar_avisos(self,veintiocho,sesenta_y_seis):
        '''Normaliza los avisos y devuelve varios DF'''
        
        #lee los archivos de los avisos y los junta (66 y 28)
        df28= pd.read_excel(r'{}'.format(veintiocho))
        df66= pd.read_excel(r'{}'.format(sesenta_y_seis))
        df66 = df66[['Aviso','Txt. cód. med.']]
        df28 = pd.merge(df28,df66,on='Aviso',how='left')

        #Suma 3 columnas a partir de las herramientas normalizadoras
        df28.rename(columns = {'Txt. cód. med.':'Prestacion'},inplace = True)
        df28['Barrio'] = df28['Emplazamiento'].map(self.herramientas.barrios).fillna('Ver')
        df28['Status simplificado'] = df28['Status usuario'].map(self.herramientas.status_simplificado)
        df28['Prestacion Simplificada']= df28['Prestacion'].map(self.herramientas.prestaciones_avisos)

        #crea un df de los n° de avisos relevados 
        df_relevamientos = Archivos_drive(self.credenciales,"R11","Respuestas de formulario 1")#self.credenciales
        relevamientos = df_relevamientos.abrir_archivo()
        avisos_relevados = list(set(relevamientos['Numero de aviso']))

        #divide en diferentes Df 
        avisos_ppu = df28.loc[df28['Grupo planif.']== 'ARP']
        avisos_general = df28.copy()
        avisos_sin_relevamiento = avisos_general.loc[~avisos_general['Aviso'].isin(avisos_relevados)]
        avisos_sin_relevamiento = avisos_sin_relevamiento.loc[~avisos_general['Status usuario'].isin(self.herramientas.sacar_status)]
        
        return [avisos_general,avisos_sin_relevamiento,avisos_ppu]
