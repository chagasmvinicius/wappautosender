from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
from PIL import Image
import io
from AppKit import NSPasteboard, NSPasteboardTypePNG
import pyperclip

keep_running = True
time_for_wapp_authentication = 3
time_general_to_sleep = 3
message = 'Está chegando o evento comemorativo de 15 anos do Proqualis / ENSP / Fiocruz!\n\n🟠 Data: 17/07/2024\n🟠 Evento presencial: Auditório térreo da ENSP - Fiocruz - Campus Manguinhos\n🟠 Transmissão on-line via canal do YouTube do Proqualis\n🟠 Certificado de participação\n\n➡ Conheça nossa programação e inscreva-se: https://bit.ly/eventoproqualis\n\n➡ Compartilhe e convide seus amigos!'

contacts_workbook = openpyxl.load_workbook('./files/contatos.xlsx')
contacts = contacts_workbook.active

# Abrir o chrome e autenticar o WhatsApp Web
driver = webdriver.Chrome()
driver.execute_script('location.replace("https://web.whatsapp.com/")')
time.sleep(time_for_wapp_authentication)

def copyImageFile():
    img = Image.open('./files/image.png') 
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_data = img_byte_arr.getvalue() 
    pasteboard = NSPasteboard.generalPasteboard()
    pasteboard.clearContents()
    pasteboard.setData_forType_(img_data, NSPasteboardTypePNG)

while keep_running:

    # Ler cada linha da planilha e salvar a informação do nome do contato e do número de telefone do contato
    for row in contacts.iter_rows(min_row = 2, values_only = True):
        name, contact = row

        search_input = driver.find_element(By.CLASS_NAME, 'selectable-text')
        time.sleep(time_general_to_sleep)

        search_input.click()
        time.sleep(time_general_to_sleep)

        search_input.send_keys(contact)
        time.sleep(time_general_to_sleep)

        search_input.send_keys(Keys.ENTER)
        time.sleep(time_general_to_sleep)

        message_input = driver.find_elements(By.CLASS_NAME, 'selectable-text')[-1]
        time.sleep(time_general_to_sleep)

        message_input.click()
        time.sleep(time_general_to_sleep)

        copyImageFile()
        message_input.send_keys(Keys.COMMAND, 'v')
        time.sleep(time_general_to_sleep)

        image_message_input = driver.find_elements(By.CLASS_NAME, 'selectable-text')[0]
        time.sleep(time_general_to_sleep)

        image_message_input.click()
        time.sleep(time_general_to_sleep)

        pyperclip.copy(message)
        image_message_input.send_keys(Keys.COMMAND, 'v')
        time.sleep(time_general_to_sleep)

        image_message_input.send_keys(Keys.ENTER)
        time.sleep(time_general_to_sleep)

    keep_running = False
driver.close()
