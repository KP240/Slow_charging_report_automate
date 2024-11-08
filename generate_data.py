import os
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import traceback
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import shutil  # For zipping the folder
import textwrap
# Database connection details
host_master = "34.100.223.97"
port = "5432"
database_master = "master_prod"
user = "postgres"
password = "theimm0rtaL"

# Create SQLAlchemy engine connections
engine_master = create_engine(f"postgresql://{user}:{password}@{host_master}:{port}/{database_master}")

# Function to get the start and end date of the previous week
def get_previous_week_dates():
    today = datetime.today()
    start_of_week = today - timedelta(days=today.weekday())  # Monday of the current week
    start_of_previous_week = start_of_week - timedelta(weeks=1)
    end_of_previous_week = start_of_week - timedelta(days=1)  # Sunday of the previous week

    return start_of_previous_week.date(), end_of_previous_week.date()

# Get the previous week's start and end date
previous_week_start, previous_week_end = get_previous_week_dates()

# Convert the dates to string format (YYYY-MM-DD) for the SQL query
previous_week_start_str = previous_week_start.strftime('%Y-%m-%d')
previous_week_end_str = previous_week_end.strftime('%Y-%m-%d')

# SQL query to fetch raw data for the previous week
query_master = f"""
SELECT
    vehicle_num,
    date,
    data_provider,
    device_type,
    site,
    city,
    start_block ->> 'soc' AS start_soc,
    end_block ->> 'soc' AS end_soc,
    charge_meta ->> 'charge_type' AS charge_type
FROM
    public.bq_charge_report
WHERE
    date BETWEEN '{previous_week_start_str}' AND '{previous_week_end_str}'
    AND charge_meta ->> 'charge_type' IN('slow','unknown')
"""

# Query to summarize the data for the previous week
query_master_1 = f"""
WITH cte AS (
    SELECT
        vehicle_num,
        date,
        data_provider,
        device_type,
        site,
        city,
        charge_meta ->> 'charge_type' AS charge_type,
        start_block ->> 'soc' AS start_soc,
        end_block ->> 'soc' AS end_soc
    FROM
        public.bq_charge_report
    WHERE
        date BETWEEN '{previous_week_start_str}' AND '{previous_week_end_str}'
        AND charge_meta ->> 'charge_type' IN('slow','unknown')
)
SELECT
    vehicle_num,
    date,
    data_provider,
    device_type,
    site,
    city,
    charge_type,
    start_soc,
    end_soc,
   CASE
       WHEN CAST(NULLIF(end_soc, '') AS NUMERIC) = 100 THEN 'Yes'
	   when CAST(NULLIF(end_soc, '') AS NUMERIC) = 100 and
	   CAST(NULLIF(start_soc, '') AS NUMERIC) IS NULL then 'No'
	   else 'No'
    END AS Slow_Charge_has_Completed_100_Percent,
	case when charge_type  = 'slow' then 'Done'
	when charge_type = 'unknown' then 'Data Not Found'
	else 'Not Done'
	end as Slow_Charge_Done_Not_Done,
	case when CAST(NULLIF(end_soc, '') AS NUMERIC) = 100 and charge_type  = 'slow'
	then 'Completed'
	 when CAST(NULLIF(end_soc, '') AS NUMERIC) <= 100 and
	   CAST(NULLIF(start_soc, '') AS NUMERIC) IS NUll and charge_type  = 'unknown'
	   then 'Charging Data is Not Found'
	else '100% Not Completed' end as status
FROM
    cte
"""

# Load the data into pandas DataFrames for the previous week
try:
    with engine_master.connect() as connection:
        # Load raw data
        df_raw = pd.read_sql_query(text(query_master), connection)
        print("Raw data for the previous week loaded successfully.")

        # Load summary data using text()
        df_slow_charge = pd.read_sql_query(text(query_master_1), connection)
        print("Slow charge summary data for the previous week loaded successfully.")

        # Remove duplicates from df_slow_charge and select only the required columns
        df_slow_charge = df_slow_charge.drop_duplicates(subset=['vehicle_num', 'city', 'site', 'slow_charge_has_completed_100_percent', 'slow_charge_done_not_done', 'status'])

        # Select only the required columns
        df_slow_charge = df_slow_charge[['vehicle_num', 'site', 'city', 'slow_charge_has_completed_100_percent', 'slow_charge_done_not_done', 'status']]

except Exception as e:
    print("Error loading data:")
    traceback.print_exc()

# Grouping by 'city', 'site', 'slow_charge_done_not_done', and 'status' to count 'status' occurrences
df_summary = df_slow_charge.groupby(['city', 'site', 'slow_charge_done_not_done', 'status']).agg(
    Count_of_Completed_100_or_not=('vehicle_num', 'nunique')
).reset_index()

# Save the summary data to a CSV file
df_summary.to_csv(f"summary_week_{previous_week_start_str}_to_{previous_week_end_str}.csv", index=False)

# Create main 'cities' folder if it doesn't exist
os.makedirs('cities', exist_ok=True)

# Save the data for each city into separate folders for the previous week
for city in df_raw['city'].unique():
    city_folder = f'cities/{city}'
    os.makedirs(city_folder, exist_ok=True)

    # Filter data by city
    df_raw_city = df_raw[df_raw['city'] == city]
    df_slow_charge_city = df_slow_charge[df_slow_charge['city'] == city]
    df_summary_city = df_summary[df_summary['city'] == city]

    # Define the output file path for the city
    output_file = f"{city_folder}/Charging-Data-{previous_week_start_str}_to_{previous_week_end_str}.xlsx"

    # Save to Excel with each DataFrame on a separate sheet
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_raw_city.to_excel(writer, sheet_name='Raw Data', index=False)
        df_slow_charge_city.to_excel(writer, sheet_name='Slow Charge Data', index=False)
        df_summary_city.to_excel(writer, sheet_name='Summary', index=False)

    print(f"Data successfully written for city: {city} in {output_file}")

    # Create a snapshot of the summary DataFrame and save it as a PNG file
    # Dynamically adjust figure size based on the number of rows and columns in the DataFrame
    num_rows, num_cols = df_summary_city.shape
    fig_width = num_cols * 2  # Adjust width based on number of columns
    fig_height = num_rows * 0.2  # Adjust height based on number of rows

    plt.figure(figsize=(fig_width, fig_height))
    plt.axis('off')  # Turn off the axis

# Create a table from the DataFrame
    table = plt.table(cellText=df_summary_city.values, colLabels=df_summary_city.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)  # Allow the font size to be set manually
    table.set_fontsize(10)
    table.scale(1.2, 1.2)  # Scale the table for better visibility

    # Save the summary table as a PNG file
    summary_png_file = f"{city_folder}/Summary-{city}-{previous_week_start_str}_to_{previous_week_end_str}.png"
    plt.tight_layout()  # Adjust layout to prevent clipping of labels
    plt.savefig(summary_png_file)
    plt.close()  # Close the plot to free memory

    print(f"Summary snapshot saved for city: {city} in {summary_png_file}")
