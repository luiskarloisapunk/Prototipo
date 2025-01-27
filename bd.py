import gspread
from google.oauth2.service_account import Credentials

# Define el alcance (scope) para acceder a Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('speedy-solstice-434503-m6-161d71670782.json', scopes=SCOPES)
client = gspread.authorize(creds)

matrizMaterias=[]

#Hacer una funcion para extraer los datos de los google sheets y almacenarlos en matrices o diccionarios, 
# y otra para obtener la materia actual y el tutor correspondiente al alumno detectado

#guardarlo en un arreglo?
def obtenerMaterias(x):
    try:
        horarioID = '1GFUgzdS-nKockbae-v5upwNBLaYEY6LhwVSiLAlfrIw'
        spreadsheet = client.open_by_key(horarioID)
    
    
        sheet = spreadsheet.worksheet('6AMP')
        data = sheet.get_all_records() 
        matrizMaterias = [row for row in data]

        materia= matrizMaterias[x-1]["MATERIA"]
        docente= matrizMaterias[x-1]["DOCENTE"]
        return materia, docente
        
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return "(No detectado)"    

#Incompleto
def obtenerTutor():
    try:
        tutorID = '1GrPbYbjPwrsLT6d4Y98UyI9pMo0vIg6GR6T6OsxJOto'
        spreadsheet = client.open_by_key(tutorID)
    
    
        sheet = spreadsheet.worksheet('6AMP')
        data = sheet.get_all_records() 
        matrizTutores = [row for row in data]
        print(matrizTutores)
        
        return materia, docente
        
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return "(No detectado)" 
    
    
materiaActual, docenteActual =obtenerMaterias(1)
print(materiaActual+" impartida por "+docenteActual)
obtenerTutor()    