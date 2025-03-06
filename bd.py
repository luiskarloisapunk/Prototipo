import gspread
import json
from google.oauth2.service_account import Credentials
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('speedy-solstice-434503-m6-161d71670782.json', scopes=SCOPES)
client = gspread.authorize(creds)

matrizMaterias=[]
matrizTutores=[]
dias=["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES"]

#Hacer una funcion para extraer los datos de los google sheets y almacenarlos en matrices o diccionarios, 
# y otra para obtener la materia actual y el tutor correspondiente al alumno detectado

#guardarlo en un arreglo?
def obtenerDatos():
    global matrizMaterias, matrizTutores 
    try:
        horarioID = '1GFUgzdS-nKockbae-v5upwNBLaYEY6LhwVSiLAlfrIw'
        spreadsheet = client.open_by_key(horarioID)
    
        
        sheet = spreadsheet.worksheet(dias[datetime.now().weekday()])
        data = sheet.get_all_records() 
        matrizMaterias = [row for row in data]
        
        
        tutorID = '1GrPbYbjPwrsLT6d4Y98UyI9pMo0vIg6GR6T6OsxJOto'
        spreadsheet = client.open_by_key(tutorID)
    
    
        sheet = spreadsheet.worksheet("6AMP")
        data = sheet.get_all_records() 
        matrizTutores = [row for row in data]
        
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return "(No detectado)"    

#Incompleto
def buscarTutor(name):
    for y in range(len(matrizTutores)):
        if matrizTutores[y]["ALUMNO"].lower() == name.lower():
            if matrizTutores[y]["TUTOR"] != "":

                return matrizTutores[y]["TUTOR"], matrizTutores[y]["NUMERO"]
    return None
x=1
   
#obtenerDatos()

def guardarDatos():
        with open("cacheM.json","w") as f:
                json.dump(matrizMaterias,f,indent=4)
                
        with open("cacheT.json","w") as f:
                json.dump(matrizTutores,f,indent=4)        
                
def cargarDatos():
        global matrizMaterias,matrizTutores
        with open("cacheM.json","r") as f:
                matrizMaterias=json.load(f)
                
        with open("cacheT.json","r") as f:
                matrizTutores=json.load(f)        
                
                
                             
#tutor,numero=buscarTutor("Luis Karlo")
#print(tutor)
#print("521"+str(numero))

if len(matrizMaterias)==0:
        cargarMaterias()


if matrizMaterias[x-1]["MATERIA"]!="RECESO":
        materiaActual= matrizMaterias[x-1]["MATERIA"]
        docenteActual= matrizMaterias[x-1]["DOCENTE"]
        print(materiaActual.title()+" impartida por "+docenteActual.title())
else: print("Es receso")         
