

class Herramientas_normalizadoras():
    """Esta clase esta pensada para contener listas y funciones que nos sirvan a la hora de normalizar archivos"""
    def __init__(self):
        """Define listas como atributos para usar en determinados casos"""
        self.corte_raices = {
            'Corte de Raices Superficial con poda aérea':'AR-CR06',
       'Corte de Raices Profunda sin poda aérea':'AR-CR07',
       'Corte de Raices Superficial sin poda aérea':'AR-CR05',
       'Corte de Raices Profunda con poda aérea':'AR-CR08',
       'Corte de Raices Superficial y profunda con poda aérea':'AR-CR10',
       'Corte de Raices Superficial y profunda sin poda aérea':'AR-CR09',
        'Agrandar Plantera':'Agrandar Plantera'
        }
        
        self.extracciones ={
            'Extracción de Árbol':'AR-EA','Extracción de Tocón':'AR-RC',
       'Extracción de Cepa':'AR-RC'
        }
        self.sacar_status = [
            'CANC','IM01','IM02','IM03',
                'IM04','IM05','SERV','FREN',
                'TERC','OTRA','OPER','PROG',
                'REOK'
        ]
        self.prestaciones ={
                'AR-P107':'Poda','AR-P207':'Poda','AR-P307':'Poda',
                'AR-P108':'Poda','AR-P208':'Poda','AR-P308':'Poda',
                'AR-P109':'Poda','AR-P209':'Poda','AR-P309':'Poda',
                'AR-P110':'Poda','AR-P210':'Poda','AR-P310':'Poda',
                'AR-P111':'Poda','AR-P211':'Poda','AR-P311':'Poda',
                'AR-P112':'Poda','AR-P212':'Poda','AR-P312':'Poda',
                'AR-P401':'Poda','AR-PP01':'Poda Puntual','AR-PP02':'Poda Puntual',
                'AR-EA07':'Extraccion','AR-EA08':'Extraccion','AR-EA09':'Extraccion',
                'AR-EA10':'Extraccion','AR-EA11':'Extraccion','AR-EA12':'Extraccion',
                'AR-EA13':'Extraccion','AR-EA14':'Extraccion','AR-RC09':'Retiro de Cepa',
                'AR-RC10':'Retiro de Cepa','AR-RC11':'Retiro de Cepa','AR-RC12':'Retiro de Cepa',
                'AR-RC13':'Retiro de Cepa','AR-RC14':'Retiro de Cepa','AR-CR05':'Corte de Raiz',
                'AR-CR06':'Corte de Raiz','AR-CR07':'Corte de Raiz','AR-CR08':'Corte de Raiz',
                'AR-CR09':'Corte de Raiz','AR-CR10':'Corte de Raiz','AR-PL03':'Plantacion',
                'AR-RP01':'Retiro de Poda','AR-RP02':'Retiro de Poda','AR-RP03':'Retiro de Poda',
                'AR-RV35':'Vereda','AR-RV36':'Vereda','AR-RV37':'Plantera'
        }
        
        self.meses={
            'Enero':{'inicio extremo':'01.01.2022','fin extremo':'30.01.2022','campo clasi':'01/2022'},
            'Febrero':{'inicio extremo':'01.02.2022','fin extremo':'28.02.2022','campo clasi':'02/2022'},
            'Marzo':{'inicio extremo':'01.03.2022','fin extremo':'31.01.2022','campo clasi':'03/2022'},
            'Abril':{'inicio extremo':'01.04.2022','fin extremo':'30.01.2022','campo clasi':'04/2022'},
            'Mayo':{'inicio extremo':'01.05.2022','fin extremo':'31.01.2022','campo clasi':'05/2022'},
            'Junio':{'inicio extremo':'01.06.2022','fin extremo':'30.01.2022','campo clasi':'06/2022'},
            'Julio':{'inicio extremo':'01.07.2022','fin extremo':'31.01.2022','campo clasi':'07/2022'},
            'Agosto':{'inicio extremo':'01.08.2022','fin extremo':'31.01.2022','campo clasi':'08/2022'},
            'Septiembre':{'inicio extremo':'01.09.2022','fin extremo':'30.01.2022','campo clasi':'09/2022'},
            'Octubre':{'inicio extremo':'01.10.2022','fin extremo':'31.01.2022','campo clasi':'10/2022'},
            'Noviembre':{'inicio extremo':'01.11.2022','fin extremo':'30.01.2022','campo clasi':'11/2022'},
            'Diciembre':{'inicio extremo':'01.12.2022','fin extremo':'31.01.2022','campo clasi':'12/2022'}
        }


    def podas(self,x):
        if x == 'Limpieza + 1 tipo de poda':
            return 'AR-P1'

        elif x == 'Limpieza + 2 tipos de podas':
            return 'AR-P2'

        elif x ==  'Limpieza + 3 tipos de podas':
            return 'AR-P3'

        elif x == 'Poda de reducción en grandes ejemplares > 20 m de altura':
            return 'AR-P401'

        else:
            return 'ERROR'

    def altura_podas(self,x):

        if x <= 4 :
            return '07'
        elif x > 4 and x <=8 :
            return '08'
        elif x > 8 and x <=12:
            return '09'
        elif x > 12 and x <=16:
            return '10'
        elif x > 16 and x <=20:
            return '11'
        elif x > 20:
            return '12'
        else:
            return 'ERROR'

    def dap(self,x):

        if x <= 5 :
            return '07'
        elif x > 5 and x <=10 :
            return '08'
        elif x > 10 and x <=20:
            return '09'
        elif x > 20 and x <=40:
            return '10'
        elif x > 40 and x <=60:
            return '11'
        elif x > 60 and x <=80:
            return '12'
        elif x > 80 and x <=100:
            return '13'
        elif x > 100:
            return '14'
        else:
            return 'ERROR'