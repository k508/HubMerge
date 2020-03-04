# HubMerge
Combine Deals and Contacts exports from Hubspot into a single useable document that can be filtered and mail merged using a docx template.

# Instructions
* Export required information from Deals section on Hubspot, rename the file ```deal.csv``` and save it in the application directory. Select all properties for the export as we are going to filter out what we don't need anyway. (Filter by ```Deal Stage``` & ```Appointment Date```)
* Export required contact information relating to the Deals and save it as ```contacts.csv``` inside the application directory. Again, export all properties. (Filter by ```Next Activity Date```)
* Save your word document template with the matching merge fields as ```template.docx``` inside the application directory.

* Run the application.

# Modifying the script for your own use

You will likely need to adjust the fields being filtered to match what you need in your own template which you can do by editing the ```merge_fields``` dictionary (see below).

```
merge_fields = ['Phone Number', 
'Appointment State/Region', 
'Appointment Date', 
'Appointment Street Address', 
'Consultant', 
'Appointment City', 
'Conference Venues', 
'Pms ID', 
'Appointment Postal Code', 
'Email', 
'Conference Date', 
'Appointment Date Time', 
'Deal Name', 
'Phone 4', 
'Phone 2', 
'Phone 3']
```

Adjust the ```document.merge``` function to match the merge fields to their respective names inside of your template document:
```
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
```
## Updates
- Modified the script to delete the temporary csv files after they are no longer required.
- Script will create the output directory if one doesn't exist instead of crashing.
- Added check for ```merge_output``` directory. If it exists it will delete the directory contents before file merge, otherwise creates directory.
- Added progress statements at each point to help correct end user processes.
- Added ```close_app``` function so the end user can make sure everything was executed correctly.
- Updated ```requirements.txt```. Can now be installed by running ```pip install -r requirements.txt``` in the root directory.
- Fixed bug with function requiring raw input.
