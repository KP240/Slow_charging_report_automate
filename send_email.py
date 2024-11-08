import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import datetime

# SMTP Configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
username = "kartik@project-lithium.com"
password = "lpolrrwyvnffyynv"

# Calculate Previous Week's Dates
today = datetime.date.today()
start_of_week = today - datetime.timedelta(days=today.weekday())
previous_week_end = start_of_week - datetime.timedelta(days=1)
previous_week_start = previous_week_end - datetime.timedelta(days=6)

previous_week_start_str = previous_week_start.strftime('%Y-%m-%d')
previous_week_end_str = previous_week_end.strftime('%Y-%m-%d')

# Recipient details
city_recipients = {
    'blr': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com']},
    'ncr': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com'], 'sub_cities': ['ggn', 'noida']},
    'pnq': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com']},
    'mum': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com']},
    'chn': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com']},
    'hyd': {'to': ['kas624988@gmail.com'], 'cc': ['kartikpande12@gmail.com']},
}

base_path = "./cities"

def send_email(city, sub_city=None, to_list=None, cc_list=None, start_date=None, end_date=None):
    msg = MIMEMultipart('related')
    msg['From'] = username
    msg['To'] = ', '.join(to_list)
    msg['Cc'] = ', '.join(cc_list)
    subject = f"Repetitive Over Speeding Offenders for {city.upper()}"
    if sub_city:
        subject += f"-{sub_city.upper()}"
    msg['Subject'] = subject

    body_html = f"""
    <html>
        <body>
            <p>Dear City Ops Team,</p>
            <p>Below are the top 5 repeat over speeding offenders for this week <strong>({start_date} to {end_date})</strong>.
            Please counsel the drivers accordingly.</p>
            <p>Attached below are the data and image for <strong>{city.upper()}{'-' + sub_city.upper() if sub_city else ''}</strong>.</p>
            <p><a href="cid:{city}{'-' + sub_city if sub_city else ''}_image" target="_blank">
                <img src="cid:{city}{'-' + sub_city if sub_city else ''}_image" alt="City Image" style="width:900px; height:auto;" /></a></p>
            <p>Best regards,<br>Kartik Pandey</p>
        </body>
    </html>
    """
    msg.attach(MIMEText(body_html, 'html'))

    if sub_city:
        city_folder = os.path.join(base_path, f"{city}-{sub_city}")
    else:
        city_folder = os.path.join(base_path, city)

    # Log absolute paths for debugging
    data_file = os.path.join(city_folder, f"Charging-Data-{previous_week_start_str}_to_{previous_week_end_str}.xlsx")
    image_file = os.path.join(city_folder, f"Summary-{city}{'-' + sub_city if sub_city else ''}-{previous_week_start_str}_to_{previous_week_end_str}.png")
    print(f"Data file path: {os.path.abspath(data_file)}")
    print(f"Image file path: {os.path.abspath(image_file)}")

    if os.path.exists(data_file):
        attach_file(msg, data_file)
    else:
        print(f"Data file for {city}{'-' + sub_city if sub_city else ''} not found at {data_file}.")

    if os.path.exists(image_file):
        with open(image_file, "rb") as img:
            mime_image = MIMEImage(img.read())
            mime_image.add_header('Content-ID', f"<{city}{'-' + sub_city if sub_city else ''}_image>")
            msg.attach(mime_image)
    else:
        print(f"Image file for {city}{'-' + sub_city if sub_city else ''} not found at {image_file}.")

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.sendmail(username, to_list + cc_list, msg.as_string())
    print(f"Email sent to {city.upper()}{'-' + sub_city.upper() if sub_city else ''} team.")

def attach_file(msg, filepath):
    filename = os.path.basename(filepath)
    attachment = MIMEBase('application', 'octet-stream')
    with open(filepath, "rb") as f:
        attachment.set_payload(f.read())
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename={filename}')
    msg.attach(attachment)

for city, recipients in city_recipients.items():
    to_list = recipients['to']
    cc_list = recipients['cc']
    if 'sub_cities' in recipients:
        for sub_city in recipients['sub_cities']:
            send_email(city, sub_city=sub_city, to_list=to_list, cc_list=cc_list,
                       start_date=previous_week_start_str, end_date=previous_week_end_str)
    else:
        send_email(city, to_list=to_list, cc_list=cc_list,
                   start_date=previous_week_start_str, end_date=previous_week_end_str)
