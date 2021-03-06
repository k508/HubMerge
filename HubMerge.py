import csv
import glob
import os
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd
from mailmerge import MailMerge

from functions import close_app, print_header

print_header()

root = Tk()
root.withdraw()

contact_csv = askopenfilename(initialdir="self.__base_filepath", title="Select Contacts CSV", filetypes=(
    ("CSV Files", "*.csv"), ("All Files", "*")))
print("You have selected", contact_csv)
deal_csv = askopenfilename(initialdir="self.__base_filepath", title="Select Deals CSV", filetypes=(
    ("CSV Files", "*.csv"), ("All Files", "*")))
print("YOu have selected", deal_csv)

a = pd.read_csv(contact_csv)
b = pd.read_csv(deal_csv)
print("Selecting Files...")

save_dir = 'merge_output/'

# If merge_output exists it deletes all files inside it, otherwise creates merge_output
if os.path.exists(save_dir):
    print("merge_output directory exists")
    _files = glob.glob('merge_output/*')
    for f in _files:
        os.remove(f)
        print("Deleting old files..")
else:
    os.makedirs(save_dir)
    print("merge_output doesn't exist. Creating directory.")

b.rename(columns={'Associated Contact IDs': 'Contact ID'}, inplace=True)

# Combine and filter csvs
print("Combining CSVs based on Contact ID")
merged = a.merge(b, on="Contact ID")
merged.to_csv("output.csv", index=False)

merge_fields = ['Phone Number', 'Appointment State/Region', 'Appointment Date', 'Appointment Street Address', 'Consultant', 'Appointment City',
                'Conference Venues', 'Pms ID', 'Appointment Postal Code', 'Email', 'Conference Date', 'Appointment Date Time', 'Deal Name', 'Phone 4', 'Phone 2', 'Phone 3', 'Consultant Note']

filtered_csv = pd.read_csv("output.csv", usecols=merge_fields)
print("Filtering Merge Fields..")
filtered_csv.to_csv("filtered_csv.csv", index=False)

# Start Mail Merge
print("Selecting template file.")
template = 'In-Home-Template.docx'

document = MailMerge(template)

print("Starting Mail Merge..")
with open('filtered_csv.csv') as file:
    reader = csv.reader(file, delimiter=',')
    next(reader)
    for merge_fields in reader:
        if merge_fields[11] != '':
            # Appointment Date Format to dd-mm-yyyy
            appointment_datetime_object = datetime.strptime(
                merge_fields[11], '%Y-%m-%d')
            appointment_datetime_format = datetime.strftime(
                appointment_datetime_object, '%d-%m-%Y')
        else:
            appointment_datetime_format = merge_fields[11]
        if merge_fields[5] != '':
            # Conference Date Format to dd-mm-yyyy
            conf_datetime_object = datetime.strptime(
                merge_fields[5], '%Y-%m-%d')
            conf_datetime_format = datetime.strftime(
                conf_datetime_object, '%d-%m-%Y')
        else:
            conf_datetime_format = merge_fields[5]
        #Add zero to number if excel stripped it
        raw_number = merge_fields[4]
        if len(raw_number) == 9:
            n = 1 # number of zeros to add
            res = raw_number.rjust(n + len(raw_number), '0')
            merge_fields[4] = res
        
        Deal_Name = merge_fields[13]
        document = MailMerge(template)
        document.merge(
            # Edit Merge Fields:
            # Template_Merge_Field = merge_fields[location]
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
            Consultant_Note=merge_fields[12],
            Deal_Name=Deal_Name,
            Appointment_Street_Address=merge_fields[14],
            Consultant=merge_fields[15],
            Appointment_Date_Time=merge_fields[16],
            )

        save_path = os.path.join(save_dir, f'In-Home-{Deal_Name}.docx')

        document.write(save_path)
        print("Created", Deal_Name)

# Delete temporary CSVs
print("Cleaning up temporary files..")
os.remove('filtered_csv.csv')
os.remove('output.csv')
os.remove(contact_csv)
os.remove(deal_csv)

close_app()
