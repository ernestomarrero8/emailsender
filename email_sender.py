import openpyxl
import win32com.client as win32
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import mammoth
import base64

def send_emails():
    # Obtener el archivo de Excel con los destinatarios
    excel_file = askopenfilename(title="Seleccionar archivo de Excel", filetypes=[("Excel files", "*.xlsx")])
    workbook = openpyxl.load_workbook(excel_file)
    worksheet = workbook.active

    # Obtener el archivo de Word con el contenido del correo
    word_file = askopenfilename(title="Seleccionar archivo de Word", filetypes=[("Word files", "*.docx")])

    # Convertir el archivo de Word a formato HTML
    with open(word_file, "rb") as file:
        result = mammoth.convert_to_html(file)
        html = result.value

    # Cargar la imagen en base64
    imagen_path = filedialog.askopenfilename(
        title="Seleccionar imagen",
        filetypes=(("Image files", "*.jpg;*.jpeg;*.png"), ("All files", "*.*"))
    )
    with open(imagen_path, "rb") as file:
        image_data = file.read()
        image_base64 = base64.b64encode(image_data).decode()

    # Enviar el correo a cada destinatario del archivo de Excel
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        recipient_email = row[0]
        recipient_first_name = row[1]

        outlook = win32.Dispatch('outlook.application')
        message = outlook.CreateItem(0)
        message.To = recipient_email
        message.Subject = "Predice en Recursos Humanos con el Software Codify!🦾🧠"

        # Añadir saludo y nombre del destinatario al correo electrónico
        message.HTMLBody = f'Hola {recipient_first_name},<br><br>'

        # Insertar contenido del Word en el correo electrónico
        message.HTMLBody += html.replace('\n', '<br>')

        # Añadir imagen al correo electrónico
        attachment = message.Attachments.Add(imagen_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "imagen.jpg")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3710001F", "image/jpeg")

        html_body = f'<html><body>{message.HTMLBody}<br><img src="cid:{attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")}" /></body></html>'
        message.HTMLBody = html_body

        message.Send()

    print("Correos enviados exitosamente.")

if __name__ == "__main__":
    send_emails()
