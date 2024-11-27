import threading
import cv2
import os
import time
from deepface import DeepFace

# Configuración de la cámara
cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)


# Carga del clasificador de rostros
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")


reference_images = {}             #Arreglo de 'conocidos'
reference_folder = "./conocidos"  # Carpeta con imágenes de referencia, recuerden que cada archivo debe tener el nombre del 'conocido'

for file in os.listdir(reference_folder):
    if file.endswith(('.jpg', '.png', '.jpeg')):
        name = os.path.splitext(file)[0]  # Extrae el nombre del archivo sin extensión
        reference_images[name] = cv2.imread(os.path.join(reference_folder, file))


detected_name = ""  
counter = 0         # Contador de fotogramas
last_detected_time = None


def check_face(frame, roi):
    global detected_name, last_detected_time
    try:
        for name, ref_img in reference_images.items():
            if DeepFace.verify(roi, ref_img.copy())['verified']:
                detected_name = name 
                last_detected_time = time.time()  # Almacena el nombre del archivo
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
        
        if last_detected_time and current_time - last_detected_time <= 4 and detected_name !="":
         print(detected_name)           
         cv2.putText(frame, "Bienvenido, " + detected_name, (20, 450), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (3, 192, 60) if detected_name != "" else (195, 59, 35), 2)
        else:detected_name=""

        cv2.imshow('video', frame)

    # Salir si se presiona 'q'
    key = cv2.waitKey(1)
    if key == ord('q'):
        break


cap.release()
cv2.destroyAllWindows()
