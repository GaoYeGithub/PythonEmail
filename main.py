import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import sys
from docx import Document
from docx.shared import Inches

address = os.getenv("EMAIL")
password = os.getenv("PASSWORD")

mail = smtplib.SMTP('smtp.gmail.com', 587)
mail.ehlo()
mail.starttls()
if address is None or password is None:
    print("Email address or password not found in environment variables.")
    sys.exit()
mail.login(address, password)

def get_contacts(filename):
    contacts_list = {}
    with open(filename, mode='r', encoding='utf-8') as f:
        contacts = f.read().split('\n')
        if not contacts or contacts[0] == '':
            print('Contacts file is empty.')
            sys.exit()
    for item in contacts:
        parts = item.split(', ')
        contacts_list[parts[0]] = parts[1:]
    return contacts_list

def read_message(filename):
    doc = Document(filename)
    html_body = ""
    for para in doc.paragraphs:
        html_body += f"<p>{para.text}</p>"

    images = []
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            images.append(rel.target_part)

    return html_body, images

def attach_images(msg, images):
    for i, image in enumerate(images):
        image_data = image.blob
        image_mime = MIMEImage(image_data)
        image_mime.add_header('Content-ID', f'<image{i}>')
        image_mime.add_header('Content-Disposition', 'inline', filename=f'image{i}.png')
        msg.attach(image_mime)

def create_html_with_images(html_body, images):
    for i in range(len(images)):
        html_body = html_body.replace(f"<p>{{image{i}}}</p>", f'<img src="cid:image{i}">')
    return html_body

contacts = get_contacts('contacts.txt')
html_body, images = read_message('email.docx')

html_body_with_images = create_html_with_images(html_body, images)

for contact_mail in contacts:
    msg = MIMEMultipart("related")
    msg['From'] = address
    msg['To'] = contact_mail
    msg['Subject'] = "Image"

    msg_alt = MIMEMultipart("alternative")
    msg.attach(msg_alt)

    msg_text = MIMEText("This is the alternative plain text message.")
    msg_html = MIMEText(html_body_with_images, 'html')
    msg_alt.attach(msg_text)
    msg_alt.attach(msg_html)

    attach_images(msg, images)

    mail.sendmail(address, contact_mail, msg.as_string())
    print(f"Sent Successfully to {contact_mail}!")

mail.quit()
