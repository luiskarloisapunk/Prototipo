import threading
import cv2
import os
import time
from deepface import DeepFace
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
from datetime import datetime

# Configuración de la cámara
cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)


face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")

reference_images = {}
reference_folder = "./conocidos"  # Carpeta con imágenes de referencia, recuerden que cada archivo debe tener el nombre del 'conocido'

for file in os.listdir(reference_folder):
    if file.endswith(('.jpg', '.png', '.jpeg')):
        name = os.path.splitext(file)[0]  # Extrae el nombre del archivo sin extensión
        reference_images[name] = cv2.imread(os.path.join(reference_folder, file))


detected_name = ""  
counter = 0         # Contador de fotogramas
last_detected_time = None



def get_module_by_time(entry_time):
    hour = entry_time.hour
    minute = entry_time.minute

    if (7 <= hour < 8) or (hour == 8 and minute < 50):     #7:00-7:50
        return 'Modulo 1'
    elif (8 <= hour < 9) or (hour == 9 and minute < 20):   #7:50-8:40
        return 'Modulo 2'
    elif (9 <= hour < 10) or (hour == 10 and minute < 5):  #8:40-9:20
        return 'Modulo 3'
    elif (10 <= hour < 11) or (hour == 11 and minute < 50): #10:05-10:50
        return 'Modulo 4'
    elif (11 <= hour < 12) or (hour == 12 and minute < 30):  #10:50-11:40
        return 'Modulo 5'
    elif (12 <= hour < 13) or (hour == 13 and minute < 20):  #11:40-12:30
        return 'Modulo 6'
    else:                                                    #12:30-13:20
        return 'Modulo 7'


def add_to_excel(name, entry_time):
    entry_time = datetime.fromtimestamp(entry_time)

    date_str = datetime.now().strftime('%Y-%m-%d')  
    entry_time_str = entry_time.strftime('%H:%M')  

    file_name = f"{date_str}.xlsx"
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
    else:
        wb = Workbook()
        wb.remove(wb.active)
        for i in range(1, 8):
            wb.create_sheet(f'Modulo {i}')


    module_sheet = get_module_by_time(entry_time)
    sheet = wb[module_sheet]

    if sheet.max_row == 1:
        sheet.append(['Nombre', 'Entrada']) 
        sheet.append([name, entry_time_str])
    else:
        registered_names = [row[0] for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True) if row[0]]
        if name not in registered_names:
            sheet.append([name, entry_time_str])


    wb.save(file_name)





def check_face(frame, roi):
    global detected_name, last_detected_time
    try:
        for name, ref_img in reference_images.items():
            if DeepFace.verify(roi, ref_img.copy())['verified']:
                detected_name = name    # Almacena el nombre del archivo
                last_detected_time = time.time()  
                return
        detected_name = "" 
        last_detected_time = time.time()
    except ValueError:
        detected_name = ""
        last_detected_time = time.time()
            


while True:
    ret, frame = cap.read()

    if ret:
        # Conversión a blanco y negro para mejor detección
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(50, 50))

        # Dibujo de rectángulos sobre las caras detectadas
        for (x, y, w, h) in faces:
            color = (3, 192, 60) if detected_name != "" else (255,255,255)
            cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)

        # Verifica (toma una foto) cada 40 fotogramas
        if len(faces) > 0 and counter % 40 == 0:
            for (x, y, w, h) in faces:
                roi = frame[y:y+h, x:x+w] 
                threading.Thread(target=check_face, args=(frame.copy(), roi)).start()
        counter += 1
        current_time = time.time()
        current_time2=datetime.now()
        
        if last_detected_time and current_time - last_detected_time <= 4 and detected_name !="":         
         cv2.putText(frame, "Bienvenido, " + detected_name, (20, 450), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (3, 192, 60) if detected_name != "" else (195, 59, 35), 2)
         add_to_excel(detected_name, current_time)
        else:detected_name=""

        cv2.imshow('video', frame)

    # Salir si se presiona 'q'
    key = cv2.waitKey(1)
    if key == ord('q'):
        break


cap.release()
cv2.destroyAllWindows()
