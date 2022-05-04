import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np

class Actualizar_archivos :
	def __init__(self,credenciales):
		self.scope =["https://www.googleapis.com/auth/drive"]
		self.creds = ServiceAccountCredentials.\
	   from_json_keyfile_name(credenciales,self.scope)
		self.client = gspread.authorize(self.creds)

	def abrir_archivo(self,archivo,hoja):
		sheet = self.client.open(archivo)
		worksheet = sheet.worksheet(hoja)
		df = pd.DataFrame(worksheet.get_all_records())
		return df



