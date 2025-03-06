import requests


numero_tutor = "5216671881375" 
mensaje = "Estimado tutor, su hijo ha sido registrado con una inasistencia."


url = "http://localhost:3008/v1/messages"

response = requests.post(url, json={"number": numero_tutor, "message": mensaje})


if response.status_code == 200:
    print("Mensaje enviado con éxito")
else:
    print("Error al enviar mensaje:", response.text)
