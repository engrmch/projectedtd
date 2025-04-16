import pyodbc
import pandas as pd
import xlwings as xw
import os



# SQL connection details
db_config = {
    'driver': '{MariaDB ODBC 3.1 Driver}',
    'server': '',
    'database': '',
    'user': '',
    'password': '',
    'charset': 'utf8',  # Add this if you are dealing with special characters
}

file_path = os.path.join(os.path.expanduser("~"), "Desktop", "ProjectedTD", "ntr.xlsx")

# =====================================
# SECTION: NTR INCOMPLETE DETAILS
# =====================================

print("QUERYING NTR INCOMPLETE. . . . . . . . . .")
   
# SQL command to read the data
sqlQuery = """
    SELECT
    p.NTR_ID AS 'Ref#',
    p.DOWNTIME AS 'Downtime',
    p.NODE AS 'Node',
    p.PERNODE AS 'Per Node',
    p.AMPLIFIER AS 'Amplifier',
    p.LOCATION AS 'Location',
    p.SYSTEM AS 'System',
    p.SMS AS 'SMS',
    p.PTCODE AS 'PT Code',
    p.SERVICE AS 'Service Affected',
    t.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    p.TRSTART AS 'TS Start',
    p.END_TIMEPERNODE AS 'TS End',
    p.ETR AS 'ETR',
    p.DURATION_PERNODE AS 'Duration',
    p.TROUBLEDETAILS AS 'Trouble Details',
    p.RESTOREDETAILS AS 'Restore Details',
    p.EXCLUDE AS 'Exclude',
    p.EXCLUSION_REASON AS 'Exclusion Reason',
    p.NF_NUM AS 'NF#',
    p.NF_STATUS AS 'NF Status',
    p.CREATION_DATETIME AS 'Creation Date/Time',
    p.FIRST_CALLDATE AS '1st Call Date/Time',
    p.`GROUP` AS 'Group',
    p.OIC AS 'OIC',
    p.TEAM AS 'Team',
    p.POINT_OF_ORIGIN AS 'Point of Origin',
    p.CCD_REASON AS 'CCD Reason',
    p.CATEGORY AS 'Category',
    p.CAUSE AS 'Cause',
    p.CONTROLLABILITY AS 'Controllability',
    p.SUBS_COUNT_CATV AS 'Subs Affected CATV',
    p.SUBS_COUNT_SBB AS 'Subs Affected SBB',
    p.DEVICE_AFFECTED_CATV AS 'Device Affected CATV',
    p.DEVICE_AFFECTED_SBB AS 'Device Affected SBB',
    p.DETECTEDBY AS 'Detected by',
    p.RELAYTIME AS 'Relaytime',
    p.IS_CORPO AS 'Is_corpo',
    p.CORPO_LIST AS 'Corpo List',
    p.AREA AS 'Area',
    p.ALARM_TRIGGER AS 'Alarm Trigger',
    p.ALARM_SOURCE AS 'Alarm Source',
    p.DETECTION_TIME AS 'Detection Time'
FROM reports.ntr_pernode p
JOIN noc_db.ntr_tbl t ON p.NTR_ID = t.NTR_ID -- Joining condition
WHERE p.DOWNTIME >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND p.DOWNTIME < CURRENT_DATE
  AND p.PERNODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND p.NTR_ACTIVE = 1
  AND t.NTR_STATUS = 'OK'
  AND p.END_TIMEPERNODE IN ('NULL', '-', ' ')

UNION

SELECT
    p.NTR_ID AS 'Ref#',
    p.DOWNTIME AS 'Downtime',
    p.NODE AS 'Node',
    p.PERNODE AS 'Per Node',
    p.AMPLIFIER AS 'Amplifier',
    p.LOCATION AS 'Location',
    p.SYSTEM AS 'System',
    p.SMS AS 'SMS',
    p.PTCODE AS 'PT Code',
    p.SERVICE AS 'Service Affected',
    t.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    p.TRSTART AS 'TS Start',
    p.END_TIMEPERNODE AS 'TS End',
    p.ETR AS 'ETR',
    p.DURATION_PERNODE AS 'Duration',
    p.TROUBLEDETAILS AS 'Trouble Details',
    p.RESTOREDETAILS AS 'Restore Details',
    p.EXCLUDE AS 'Exclude',
    p.EXCLUSION_REASON AS 'Exclusion Reason',
    p.NF_NUM AS 'NF#',
    p.NF_STATUS AS 'NF Status',
    p.CREATION_DATETIME AS 'Creation Date/Time',
    p.FIRST_CALLDATE AS '1st Call Date/Time',
    p.`GROUP` AS 'Group',
    p.OIC AS 'OIC',
    p.TEAM AS 'Team',
    p.POINT_OF_ORIGIN AS 'Point of Origin',
    p.CCD_REASON AS 'CCD Reason',
    p.CATEGORY AS 'Category',
    p.CAUSE AS 'Cause',
    p.CONTROLLABILITY AS 'Controllability',
    p.SUBS_COUNT_CATV AS 'Subs Affected CATV',
    p.SUBS_COUNT_SBB AS 'Subs Affected SBB',
    p.DEVICE_AFFECTED_CATV AS 'Device Affected CATV',
    p.DEVICE_AFFECTED_SBB AS 'Device Affected SBB',
    p.DETECTEDBY AS 'Detected by',
    p.RELAYTIME AS 'Relaytime',
    p.IS_CORPO AS 'Is_corpo',
    p.CORPO_LIST AS 'Corpo List',
    p.AREA AS 'Area',
    p.ALARM_TRIGGER AS 'Alarm Trigger',
    p.ALARM_SOURCE AS 'Alarm Source',
    p.DETECTION_TIME AS 'Detection Time'
FROM reports.ntr_pernode p
JOIN noc_db.ntr_tbl t ON p.NTR_ID = t.NTR_ID -- Joining condition
WHERE p.DOWNTIME >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND p.DOWNTIME < CURRENT_DATE
  AND p.PERNODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND p.NTR_ACTIVE = 1
  AND t.NTR_STATUS = 'OK'
  AND (p.DURATION_PERNODE < 0)

UNION

SELECT
    p.NTR_ID AS 'Ref#',
    p.DOWNTIME AS 'Downtime',
    p.NODE AS 'Node',
    p.PERNODE AS 'Per Node',
    p.AMPLIFIER AS 'Amplifier',
    p.LOCATION AS 'Location',
    p.SYSTEM AS 'System',
    p.SMS AS 'SMS',
    p.PTCODE AS 'PT Code',
    p.SERVICE AS 'Service Affected',
    t.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    p.TRSTART AS 'TS Start',
    p.END_TIMEPERNODE AS 'TS End',
    p.ETR AS 'ETR',
    p.DURATION_PERNODE AS 'Duration',
    p.TROUBLEDETAILS AS 'Trouble Details',
    p.RESTOREDETAILS AS 'Restore Details',
    p.EXCLUDE AS 'Exclude',
    p.EXCLUSION_REASON AS 'Exclusion Reason',
    p.NF_NUM AS 'NF#',
    p.NF_STATUS AS 'NF Status',
    p.CREATION_DATETIME AS 'Creation Date/Time',
    p.FIRST_CALLDATE AS '1st Call Date/Time',
    p.`GROUP` AS 'Group',
    p.OIC AS 'OIC',
    p.TEAM AS 'Team',
    p.POINT_OF_ORIGIN AS 'Point of Origin',
    p.CCD_REASON AS 'CCD Reason',
    p.CATEGORY AS 'Category',
    p.CAUSE AS 'Cause',
    p.CONTROLLABILITY AS 'Controllability',
    p.SUBS_COUNT_CATV AS 'Subs Affected CATV',
    p.SUBS_COUNT_SBB AS 'Subs Affected SBB',
    p.DEVICE_AFFECTED_CATV AS 'Device Affected CATV',
    p.DEVICE_AFFECTED_SBB AS 'Device Affected SBB',
    p.DETECTEDBY AS 'Detected by',
    p.RELAYTIME AS 'Relaytime',
    p.IS_CORPO AS 'Is_corpo',
    p.CORPO_LIST AS 'Corpo List',
    p.AREA AS 'Area',
    p.ALARM_TRIGGER AS 'Alarm Trigger',
    p.ALARM_SOURCE AS 'Alarm Source',
    p.DETECTION_TIME AS 'Detection Time'
FROM reports.ntr_pernode p
JOIN noc_db.ntr_tbl t ON p.NTR_ID = t.NTR_ID -- Joining condition
WHERE p.DOWNTIME >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND p.DOWNTIME < CURRENT_DATE
  AND p.PERNODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND p.NTR_ACTIVE = 1
  AND t.NTR_STATUS = 'OK'
  AND (p.DURATION_PERNODE < 0)
  AND p.CATEGORY IN ('NULL', '-', ' ')

UNION

SELECT
    p.NTR_ID AS 'Ref#',
    p.DOWNTIME AS 'Downtime',
    p.NODE AS 'Node',
    p.PERNODE AS 'Per Node',
    p.AMPLIFIER AS 'Amplifier',
    p.LOCATION AS 'Location',
    p.SYSTEM AS 'System',
    p.SMS AS 'SMS',
    p.PTCODE AS 'PT Code',
    p.SERVICE AS 'Service Affected',
    t.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    p.TRSTART AS 'TS Start',
    p.END_TIMEPERNODE AS 'TS End',
    p.ETR AS 'ETR',
    p.DURATION_PERNODE AS 'Duration',
    p.TROUBLEDETAILS AS 'Trouble Details',
    p.RESTOREDETAILS AS 'Restore Details',
    p.EXCLUDE AS 'Exclude',
    p.EXCLUSION_REASON AS 'Exclusion Reason',
    p.NF_NUM AS 'NF#',
    p.NF_STATUS AS 'NF Status',
    p.CREATION_DATETIME AS 'Creation Date/Time',
    p.FIRST_CALLDATE AS '1st Call Date/Time',
    p.`GROUP` AS 'Group',
    p.OIC AS 'OIC',
    p.TEAM AS 'Team',
    p.POINT_OF_ORIGIN AS 'Point of Origin',
    p.CCD_REASON AS 'CCD Reason',
    p.CATEGORY AS 'Category',
    p.CAUSE AS 'Cause',
    p.CONTROLLABILITY AS 'Controllability',
    p.SUBS_COUNT_CATV AS 'Subs Affected CATV',
    p.SUBS_COUNT_SBB AS 'Subs Affected SBB',
    p.DEVICE_AFFECTED_CATV AS 'Device Affected CATV',
    p.DEVICE_AFFECTED_SBB AS 'Device Affected SBB',
    p.DETECTEDBY AS 'Detected by',
    p.RELAYTIME AS 'Relaytime',
    p.IS_CORPO AS 'Is_corpo',
    p.CORPO_LIST AS 'Corpo List',
    p.AREA AS 'Area',
    p.ALARM_TRIGGER AS 'Alarm Trigger',
    p.ALARM_SOURCE AS 'Alarm Source',
    p.DETECTION_TIME AS 'Detection Time'
FROM reports.ntr_pernode p
JOIN noc_db.ntr_tbl t ON p.NTR_ID = t.NTR_ID -- Joining condition
WHERE p.DOWNTIME >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND p.DOWNTIME < CURRENT_DATE
  AND p.PERNODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND p.NTR_ACTIVE = 1
  AND t.NTR_STATUS = 'OK'
  AND (p.DURATION_PERNODE < 0)
  AND p.CAUSE IN ('NULL', '-', ' ')

ORDER BY `Downtime` DESC;
 
"""

# Getting the data from SQL into a pandas DataFrame
connection = pyodbc.connect(**db_config)
df = pd.read_sql(sql=sqlQuery, con=connection)


# Display the results in the terminal
print("Query Results:")
print(df)
 

with xw.App(visible=False) as app:
    try:
        WB = xw.Book(file_path)
 
        # Check if the 'INCOMPLETE DETAILS' sheet exists
        if 'INCOMPLETE DETAILS' in [sheet.name for sheet in WB.sheets]:
            # Clear existing contents of the 'INCOMPLETE DETAILS' sheet
            WS = WB.sheets['INCOMPLETE DETAILS']
            WS.clear()  # Clears all the contents but keeps the sheet

        else:
            # If it does not exist, add a new "INCOMPLETE DETAILS" sheet at the last position
            WS = WB.sheets.add('INCOMPLETE DETAILS', after=WB.sheets[-1])

        
        # Write the DataFrame to the "INCOMPLETE DETAILS" sheet starting from A1
        WS.range("A1").options(index=False, header=True).value = df

        # Set column width and row height
        WS.range("A1").expand("table").row_height = 12
        WS.range("A1").expand("table").column_width = 15


        # Save WB with the updated data
        WB.save(file_path)

        print("DONE EXTRACTING INCOMPLETE DETAILS.")
   
    except Exception as e:
        print(f"Failed to open workbook: {e}")


