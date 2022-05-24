import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

class Archivos_drive :
	def __init__(self,credenciales,libro, hoja):
		"""Ingresa con las credenciales (es importante tenerlas y tambien los permisos de acceso,
		de no ser asi la clase no funcionara), luego el libro y la hoja"""
		
		self.scope =["https://www.googleapis.com/auth/drive"]
		self.creds = ServiceAccountCredentials.from_json_keyfile_name(credenciales,self.scope)
		self.client = gspread.authorize(self.creds)
		self.sheet = self.client.open(libro)
		self.worksheet = self.sheet.worksheet(hoja)
		
	def abrir_archivo(self):
		""" Abre los archivos y devuelve un DF, en caso de que no tenga un titulo las columnas,
		toma la primer linea y la usa de esta manera """
		
		try:
			df = pd.DataFrame(self.worksheet.get_all_records())
		except:
			df = pd.DataFrame(self.worksheet.get_values())
			titulo = df.iloc[0]
			df = df[1:]
			df.columns = titulo
			
		return df

	def borrar_archivo(self):
		"""Borra todo lo que esta en el interior de la hoja especificada"""
		self.worksheet.clear()
		
	def subir_archivo(self,df):
		"""Sube el nombre de las columnas en la primera linea 'A1' y el resto del DF a partir del  """
		self.worksheet.update("A1",[df.columns.tolist()])
		self.worksheet.update("A2",df.values.tolist())
