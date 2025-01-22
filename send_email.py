import smtplib
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import datetime
import psycopg2

# SMTP Configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
username = "kartik@project-lithium.com"
password = "lpolrrwyvnffyynv"

# PostgreSQL Configuration for fetching vehicle count
pg_host = "34.100.223.97"
pg_port = '5432'
pg_dbname = "master_prod"
pg_user = "postgres"
pg_password = "theimm0rtaL"

# Calculate Previous Week's Dates
today = datetime.date.today()
start_of_week = today - datetime.timedelta(days=today.weekday())
previous_week_end = start_of_week - datetime.timedelta(days=1)
previous_week_start = previous_week_end - datetime.timedelta(days=6)
previous_week_start_str = previous_week_start.strftime('%Y-%m-%d')
previous_week_end_str = previous_week_end.strftime('%Y-%m-%d')

# Update city_recipients to include database query names
city_recipients = {
    'blr': {'to': ['fmat247.blr@project-lithium.com','sridhar@project-lithium.com','vasu@project-lithium.com','sridhar@project-lithium.com','fmataon.blr@project-lithium.com','Gopinath@project-lithium.com','fmatanz.blr@project-lithium.com','sathish.s@project-lithium.com','fmatgoogle.blr@project-lithium.com','fmathgs.blr@project-lithium.com','fmatkpmg.blr@project-lithium.com','sap.blr@project-lithium.com','fmatcolt.blr@project-lithium.com','fmatoptum.blr@project-lithium.com','fmatpaloalto.blr@project-lithium.com','fmatsalesforce@project-lithium.com','mukesh@project-lithium.com','fmattesco@project-lithium.com','fmatunisys.blr@project-lithium.com','fmatvmware.blr@project-lithium.com','fmatvolvo.blr@project-lithium.com','fmatwellsfargo.blr@project-lithium.com']
            , 'cc': ['nithya@project-lithium.com','niloy@project-lithium.com'], 'query_name': 'blr'},
    'ncr': {
        'to': ['raju@project-lithium.com'],
        'cc': ['nithya@project-lithium.com',],
        'sub_cities': {'ggn': 'ncr', 'noida': 'ncr'},  # Map sub-city folders to database query name
    },
    'pnq': {'to': ['prashant@project-lithium.com'], 'cc': ['nithya@project-lithium.com'], 'query_name': 'pnq'},
    'mum': {'to': ['parvez.shaikh@project-lithium.com'], 'cc': ['nithya@project-lithium.com'], 'query_name': 'mum'},
    'chn': {'to': ['aravindraj@project-lithium.com'], 'cc': ['nithya@project-lithium.com','fmatli.chn@project-lithium.com'], 'query_name': 'chn'},
    'hyd': {'to': ['karan@project-lithium.com'], 'cc': ['nithya@project-lithium.com'], 'query_name': 'hyd'},
}

# Base path for the city folders
base_path = "./cities"

def fetch_vehicle_count(city_folder):
    """
    Fetch the total number of unique vehicles for a city from PostgreSQL.
    """
    # Extract the city name from the folder path (exclude the base path)
    city_name = os.path.basename(city_folder)  # This will get 'blr', 'ncr-ggn', etc.
    
    try:
        connection = psycopg2.connect(
            host=pg_host,
            port=pg_port,
            dbname=pg_dbname,
            user=pg_user,
            password=pg_password,
        )
        cursor = connection.cursor()
        query = """
            SELECT COUNT(DISTINCT vehicle_num) AS count_vehicle
            FROM bq_charge_report
            WHERE city = %s AND date BETWEEN %s AND %s;
        """
        print(f"Executing query: {query} with params: {city_name}, {previous_week_start_str}, {previous_week_end_str}")
        cursor.execute(query, (city_name, previous_week_start_str, previous_week_end_str))
        result = cursor.fetchone()
        print(f"Query result: {result}")
        return result[0] if result and result[0] is not None else 0
    except Exception as e:
        print(f"Error fetching vehicle count for {city_folder}: {e}")
        return 0
    finally:
        if connection:
            cursor.close()
            connection.close()

def analyze_charging_data(city_folder):
    """
    Analyze charging data for a city (or sub-city) from Excel files and return summary statistics.
    """
    data_file = os.path.join(city_folder, f"Charging-Data-{previous_week_start_str}_to_{previous_week_end_str}.xlsx")

    if not os.path.exists(data_file):
        print(f"Data file for {city_folder} not found at {data_file}.")
        return 0, 0, 0, 0

    try:
        df = pd.read_excel(data_file, sheet_name="Slow Charge Data")

        if 'vehicle_num' not in df.columns or 'slow_charge_has_completed_100_percent' not in df.columns:
            print(f"Required columns missing in the data for {city_folder}")
            return 0, 0, 0, 0

        slow_charge_completed_vehicles = df[df['slow_charge_has_completed_100_percent'] == 'Yes']['vehicle_num'].unique()
        slow_charge_completed = len(slow_charge_completed_vehicles)

        incomplete_slow_charge_vehicles = df[(
            df['slow_charge_has_completed_100_percent'] == 'No') & 
            (~df['vehicle_num'].isin(slow_charge_completed_vehicles))
        ]['vehicle_num'].unique()
        incomplete_slow_charge = len(incomplete_slow_charge_vehicles)

        charging_data_not_found = df[df['status'] == 'Charging Data Not Found']['vehicle_num'].nunique()

        total_vehicles = fetch_vehicle_count(city_folder)

        return total_vehicles, slow_charge_completed, incomplete_slow_charge, charging_data_not_found

    except Exception as e:
        print(f"Error analyzing charging data for {city_folder}: {e}")
        return 0, 0, 0, 0

def attach_file(msg, filepath):
    """
    Attach file to the email message.
    """
    try:
        with open(filepath, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filepath)}')
            msg.attach(part)
    except Exception as e:
        print(f"Error attaching file {filepath}: {e}")

def attach_inline_image(msg, filepath, cid):
    """
    Attach an inline image to the email.
    """
    try:
        with open(filepath, 'rb') as img:
            img_part = MIMEImage(img.read())
            img_part.add_header('Content-ID', f'<{cid}>')
            img_part.add_header('Content-Disposition', 'inline', filename=os.path.basename(filepath))
            msg.attach(img_part)
    except Exception as e:
        print(f"Error attaching inline image {filepath}: {e}")

def send_email(city, sub_city=None, to_list=None, cc_list=None):
    """
    Send the weekly charging data report email for a city or sub-city.
    """
    # Construct the correct folder path
    if sub_city:
        city_folder = os.path.join(base_path, f"{city}-{sub_city}")
    else:
        city_folder = os.path.join(base_path, city)
    
    # Extract the city name from the folder (use last folder name if sub-city)
    city_name = os.path.basename(city_folder)

    # Fetch vehicle count using the query name
    query_name = (
        city_recipients[city]['sub_cities'].get(sub_city, city_recipients[city].get('query_name', city))
        if sub_city else city_recipients[city].get('query_name', city)
    )

    total_vehicles, completed_slow_charge, incomplete_slow_charge, charging_data_not_found = analyze_charging_data(city_folder)
    
    remaining_vehicles = total_vehicles - (completed_slow_charge + incomplete_slow_charge + charging_data_not_found)

    msg = MIMEMultipart('related')
    msg['From'] = username
    msg['To'] = ', '.join(to_list)
    msg['Cc'] = ', '.join(cc_list)
    subject = f"{city.upper()}{'-' + sub_city.upper() if sub_city else ''} - Weekly Slow Charging Report-2025"
    msg['Subject'] = subject

    body_html = f"""
<html>
    <body>
        <p>Slow Charging Report Date For {previous_week_start_str} to {previous_week_end_str}</p>
        <p>Kindly review the vehicles listed below that are mapped to {city.upper()}{'-' + sub_city.upper() if sub_city else ''} sites.</p>
        <p>Out of {total_vehicles} vehicles, only {completed_slow_charge} have completed slow charging to 100% as per MMI data, and {incomplete_slow_charge} vehicles have not completed slow charging,Remaining {remaining_vehicles} Vehicles Charging Data is not found Please ensure that all vehicles undergo slow charging on a weekly basis.</p>
        <p>Attaching the file for the detailed report. Kindly check.</p>
        <p>See the summary below:</p>
            <img src="cid:summary_image" alt="Summary Image" width="600">
        <p>Best regards,<br>Kartik Pandey</p>
    </body>
</html>
"""
    msg.attach(MIMEText(body_html, 'html'))

    # Attach the charging data file
    data_file = os.path.join(city_folder, f"Charging-Data-{previous_week_start_str}_to_{previous_week_end_str}.xlsx")
    if os.path.exists(data_file):
        attach_file(msg, data_file)

    # Attach inline image (if available)
    summary_image_path =  os.path.join(city_folder, f"Summary-{city}{'-' + sub_city if sub_city else ''}-{previous_week_start_str}_to_{previous_week_end_str}.png")  # Update this path to the actual image
    attach_inline_image(msg, summary_image_path, 'summary_image')

    try:
        # Send the email via SMTP
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(username, password)
            server.sendmail(msg['From'], to_list + cc_list, msg.as_string())
            print(f"Email sent to {to_list} and cc: {cc_list} for {city_name}")
    except Exception as e:
        print(f"Error sending email for {city_name}: {e}")

# Run the email sending for each city
for city, city_data in city_recipients.items():
    # Skip 'ncr' and process sub-cities individually
    if city == 'ncr':
        for sub_city, query_name in city_data['sub_cities'].items():
            send_email(city, sub_city=sub_city, to_list=city_data['to'], cc_list=city_data['cc'])
    else:
        send_email(city, to_list=city_data['to'], cc_list=city_data['cc'])

    # If sub-cities exist for other cities (not 'ncr'), send separate emails for each sub-city
    if 'sub_cities' in city_data and city != 'ncr':  # Skip sub-cities for 'ncr'
        for sub_city, query_name in city_data['sub_cities'].items():
            send_email(city, sub_city=sub_city, to_list=city_data['to'], cc_list=city_data['cc'])
