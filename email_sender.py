import yagmail
import os
import json
from dotenv import load_dotenv

def Send_Mail(path):
    try:
        load_dotenv('./dotenv.env')
        # Initialize yagmail connection (configure once)
        yag = yagmail.SMTP(
            user= os.getenv('EMAIL_USER'),          # Your email address
            password= os.getenv('EMAIL_PASSWORD'),  # For Gmail, use App Password if 2FA enabled
            host= os.getenv('HOST_MAIL'),       # SMTP server (change for other providers)
            port=465,                     # SSL port
            smtp_ssl=True                 # Use SSL
        )

        # Email details
        email_subject = "Revisión de Tintas Semanal"
        email_body = """
        Estimado equipo de sistemas,

        A traves del presente medio se le hace llegar el informe de tintas semanal a traves del script automatizado.


        Saludos,
        Sebastian Etchyberre
        """

        # Attachments (can be single file or list)
        pdf_attachment = path

        # Recipients
        to = os.getenv("RECIPIENT")          # Main recipient
        cc = json.loads(os.getenv('EMAIL_CC_LIST', '[]'))

        # Send email with all components
        yag.send(
            to=to,
            cc=cc,
            subject=email_subject,
            contents=email_body,
            attachments=pdf_attachment
        )

        print("Email sent successfully to main recipient and CC recipients!")
    except Exception as e:
        print("Review .env file and check if the data is correct")
    