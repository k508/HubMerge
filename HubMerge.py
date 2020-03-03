from mailmerge import MailMerge
from datetime import datetime
import os
import pandas as pd
import csv

a = pd.read_csv("contacts.csv")
b = pd.read_csv("deals.csv")

b.rename(columns={'Associated Contact IDs': 'Contact ID'}, inplace=True)

# Combine and filter csvs
merged = a.merge(b, on="Contact ID")
merged.to_csv("output.csv", index=False)

merge_fields = ['Phone Number', 'Appointment State/Region', 'Appointment Date', 'Appointment Street Address', 'Consultant', 'Appointment City',
                'Conference Venues', 'Pms ID', 'Appointment Postal Code', 'Email', 'Conference Date', 'Appointment Date Time', 'Deal Name', 'Phone 4', 'Phone 2', 'Phone 3']

filtered_csv = pd.read_csv("output.csv", usecols=merge_fields)

filtered_csv.to_csv("filtered_csv.csv", index=False)

# Start Mail Merge
template = 'template.docx'

document = MailMerge(template)


with open('filtered_csv.csv') as file:
    reader = csv.reader(file, delimiter=',')
    next(reader)
    for merge_fields in reader:
        if merge_fields[11] != '':
            appointment_datetime_object = datetime.strptime(
                merge_fields[11], '%Y-%m-%d')
            appointment_datetime_format = datetime.strftime(
                appointment_datetime_object, '%d-%m-%Y')
        else:
            appointment_datetime_format = merge_fields[11]
        if merge_fields[5] != '':
            conf_datetime_object = datetime.strptime(
                merge_fields[5], '%Y-%m-%d')
            conf_datetime_format = datetime.strftime(
                conf_datetime_object, '%d-%m-%Y')
        else:
            conf_datetime_format = merge_fields[5]
        Deal_Name = merge_fields[12]
        document = MailMerge(template)
        document.merge(
            Phone_2=merge_fields[0],
            Phone_3=merge_fields[1],
            Phone_4=merge_fields[2],
            Pms_ID=merge_fields[3],
            Phone_Number=merge_fields[4],
            Conference_Date=conf_datetime_format,
            Email=merge_fields[6],
            Conference_Venues=merge_fields[7],
            Appointment_Postal_Code=merge_fields[8],
            Appointment_StateRegion=merge_fields[9],
            Appointment_City=merge_fields[10],
            Appointment_Date=appointment_datetime_format,
            Deal_Name=merge_fields[12],
            Appointment_Street_Address=merge_fields[13],
            Consultant=merge_fields[14],
            Appointment_Date_Time=merge_fields[15],
        )

        save_dir = 'merge_output/'
        save_path = os.path.join(save_dir, f'In-Home-{Deal_Name}.docx')

        document.write(save_path)
