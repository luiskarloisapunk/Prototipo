import cv2
import time
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import os
import face_recognition
import cv2
import os
import glob
import numpy as np


cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)



ultimaDeteccion = None
known_face_encodings = []
known_face_names = []
frame_resizing = 0.25


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


def load_encoding_images(images_path):

    images_path = glob.glob(os.path.join(images_path, "*.*"))

    print("Se han encontrado {} imagenes en la base de datos. Codificando...".format(len(images_path)))

        # Store image encoding and names
    for img_path in images_path:
        img = cv2.imread(img_path)
        rgb_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

            # Get the filename only from the initial file path.
        basename = os.path.basename(img_path)
        (filename, ext) = os.path.splitext(basename)
            # Get encoding
        img_encoding = face_recognition.face_encodings(rgb_img)[0]

            # Store file name and file encoding
        known_face_encodings.append(img_encoding)
        known_face_names.append(filename)
    print("Imagenes cargadas.")

def detect_known_faces(frame):
    global ultimaDeteccion
    small_frame = cv2.resize(frame, (0, 0), fx=frame_resizing, fy=frame_resizing)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
    
    # Detectar ubicaciones y codificaciones de rostros en el fotograma
    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

    face_names = []
    for face_encoding in face_encodings:
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = ""

        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = np.argmin(face_distances)
        if matches[best_match_index]:
            name = known_face_names[best_match_index]
            ultimaDeteccion = time.time()

        face_names.append(name)

    # Asegurarse de devolver siempre algo
    if len(face_locations) == 0:
        return [], []

    face_locations = np.array(face_locations)
    face_locations = (face_locations / frame_resizing).astype(int)
    
    return face_locations, face_names



load_encoding_images("conocidos/")




#Bucle principal
while True:
    ret, frame = cap.read()

    # Detect Faces
    face_locations, face_names = detect_known_faces(frame)
    for face_loc, name in zip(face_locations, face_names):
        y1, x2, y2, x1 = face_loc[0], face_loc[1], face_loc[2], face_loc[3]
    
        actual = time.time()
        
        if  name !="":
            cv2.putText(frame, "Bienvenido, " + name, (20, 450), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (3, 192, 60) if name != "" else (195, 59, 35), 2)
            add_to_excel(name, actual)
        color = (3, 192, 60) if name != "" else (255,255,255)
        cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)

    cv2.imshow("CHECKPOINT", frame)

    key = cv2.waitKey(1)
    if key == ord("q"):
        break

cap.release()
cv2.destroyAllWindows()