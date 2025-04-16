import pyodbc
import pandas as pd
import os
import fnmatch
import getpass
import xlwings as xw



# SQL connection details
db_config = {
    'driver': '{MariaDB ODBC 3.1 Driver}',
    'server': '1',
    'database': '',
    'user': '',
    'password': '',
    'charset': 'utf8',  # Add this if you are dealing with special characters
}

file_path = os.path.join(os.path.expanduser("~"), "Desktop", "ProjectedTD", "ntr.xlsx")

# =====================================
# SECTION: NTR 60 DAYS
# =====================================

print("QUERYING NTR (60 DAYS). . . . . . . . . .")
 
   
# SQL command to read the data
sqlQuery = """
SELECT
    n.NTR_ID AS 'Ref#',
    CONCAT(DATE_FORMAT(n.NTR_DATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TIME, '%H:%i')) AS 'Downtime',
    n.NTR_NODE AS 'Node',
    s.NTR_NODE AS 'Per Node',
    n.NTR_AMPLIFIER AS 'Amplifier',
    n.NTR_LOC AS 'Location',
    n.NTR_SYSTEM AS 'System',
    n.NTR_SMS AS 'SMS',
    n.NTR_PTCODE AS 'PT Code',
    n.NTR_SERVICE AS 'Service Affected',
    n.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    CONCAT(DATE_FORMAT(n.NTR_TRDATESTART, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TRTIMESTART, '%H:%i')) AS 'TS Start',
    s.TS_END AS 'TS End',
    n.NTR_ETR AS 'ETR',
    ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) AS 'Duration',
    n.NTR_TROUBLEDETAILS AS 'Trouble Details',
    n.NTR_RESTOREDETAILS AS 'Restore Details',
    n.NTR_EXCLUDE AS 'Exclude',
    n.NTR_EXCLUDE_CAT AS 'Exclusion Reason',
    n.NTR_FAULTNUM AS 'NF#',
    n.NTR_FAULTSTS AS 'NF Status',
    CONCAT(DATE_FORMAT(n.NTR_CRDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_CRTIME, '%H:%i')) AS 'Creation Date/Time',
    CONCAT(DATE_FORMAT(n.NTR_1CDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_1CTIME, '%H:%i')) AS '1st Call Date/Time',
    n.NTR_GROUP AS 'Group',
    n.NTR_OIC AS 'OIC',
    n.NTR_TEAM AS 'Team',
    n.NTR_PORIGIN AS 'Point of Origin',
    n.NTR_CCD_REASON AS 'CCD Reason',
    n.NTR_CATEGORY AS 'Category',
    n.NTR_CAUSE AS 'Cause',
    n.NTR_CONTROLLABILITY AS 'Controllability',
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'CATV' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected CATV`,
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'SBB' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected SBB`,
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '1W') AS 'Device Affected CATV',
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '2W') AS 'Device Affected SBB',
    n.NTR_ONBOARD AS 'Detected by',
    n.NTR_RELAYTIME AS 'Relaytime',
    n.NTR_CORPO AS 'Is_corpo',
    n.NTR_CORPO_LIST AS 'Corpo List',
    n.NTR_GROUP AS 'Area',
    n.NTR_ALARM_TRIGGER AS 'Alarm Trigger',
    n.NTR_ALARM_SOURCE AS 'Alarm Source',
    n.NTR_TIMESTAMP AS 'Detection Time'
FROM noc_db.ntr_tbl n
JOIN noc_db.ntr_subs_affected_tbl s ON n.NTR_ID = s.NTR_ID
WHERE TIMESTAMP(n.NTR_DATE, n.NTR_TIME) >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND TIMESTAMP(n.NTR_DATE, n.NTR_TIME) < CURRENT_DATE
  AND s.NTR_NODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND TRIM(LOWER(n.NTR_STATUS)) NOT IN ('cancelled')
  AND s.ACTIVE = 1
GROUP BY n.NTR_ID, n.NTR_NODE, s.NTR_NODE, n.NTR_AMPLIFIER, n.NTR_LOC, n.NTR_SYSTEM, 
         n.NTR_SMS, n.NTR_PTCODE, n.NTR_SERVICE, n.NTR_STATUS, n.NTR_TRDATESTART, 
         n.NTR_TRTIMESTART, s.TS_END, n.NTR_ETR, n.NTR_TROUBLEDETAILS, n.NTR_RESTOREDETAILS, 
         n.NTR_EXCLUDE, n.NTR_EXCLUDE_CAT, n.NTR_FAULTNUM, n.NTR_FAULTSTS, n.NTR_CRDATE, 
         n.NTR_CRTIME, n.NTR_1CDATE, n.NTR_1CTIME, n.NTR_GROUP, n.NTR_OIC, n.NTR_TEAM, 
         n.NTR_PORIGIN, n.NTR_CCD_REASON, n.NTR_CATEGORY, n.NTR_CAUSE, n.NTR_CONTROLLABILITY, 
         n.NTR_ONBOARD, n.NTR_RELAYTIME, n.NTR_CORPO, n.NTR_CORPO_LIST, n.NTR_ALARM_TRIGGER, 
         n.NTR_ALARM_SOURCE, n.NTR_TIMESTAMP
ORDER BY TIMESTAMP(n.NTR_DATE, n.NTR_TIME) DESC;
 
"""
 
# Getting the data from SQL into a pandas DataFrame
connection = pyodbc.connect(**db_config)
df = pd.read_sql(sql=sqlQuery, con=connection)

# Apply condition to set Category to 'Blank' for ON-GOING or PENDING status
df.loc[(df['Status'].isin(['ON-GOING', 'PENDING'])) & (df['Category'].isnull() | (df['Category'] == '')), 'Category'] = 'Blanks'
 
# Add the '><=4' column with calculated values based on the DURATION column
df['><=4'] = (df['Duration'] < 4).map({True: "<4", False: ">=4"})


with xw.App(visible=False) as app:
    app.screen_updating = False
    app.calculation = 'manual'
    try:
        WB = xw.Book(file_path)

        # Disable Excel automatic calculation
        WB.app.calculation = 'manual'
 
        # Check if the 'NTR60' sheet exists
        if 'NTR60' in [sheet.name for sheet in WB.sheets]:
            # Clear existing contents of the 'NTR60' sheet
            WS = WB.sheets['NTR60']
            WS.clear()  # Clears all the contents but keeps the sheet

        else:
            # If it does not exist, add a new "NTR60" sheet at the last position
            WS = WB.sheets.add('NTR60', after=WB.sheets[-1])

        
        # Write the DataFrame to the "NTR60" sheet starting from A1
        WS.range("A1").options(index=False, header=True).value = df

        # Set column width and row height
        WS.range("A1").expand("table").row_height = 12
        WS.range("A1").expand("table").column_width = 15

        print("DONE EXTRACTING NTR(60 DAYS).")

        # Refresh the pivot table only in the "NTR_60" sheet
        ntr60_pivot = "NTR_60"  # The name of the pivot table sheet
        pivot_sheet = WB.sheets[ntr60_pivot]

        print("REFRESHING PIVOT TABLE IN SHEET NTR_60. DO NOT CLOSE!")

        for pivot_table in pivot_sheet.api.PivotTables():
            pivot_table.RefreshTable()

            # Define the data range explicitly
            data_range = WS.range("A1").expand("table")

            pivot_sheet.range("A3").value = '=TEXT(MIN(NTR60!B:B),"MM/DD/YYYY")&" - "&TEXT(MAX(NTR60!B:B),"MM/DD/YYYY")'
            
            # Save WB with the updated data
            WB.save(file_path)


        print("DONE REFRESHING PIVOT TABLE AT NTR_60.")
    except Exception as e:
        print(f"Failed to open workbook: {e}")

    print("APPLYING FORMULAS.")

    #DT (Hr)
    formula_cell = pivot_sheet.range('E6')  # Cell where the formula starts
    formula_cell.value = '=IF(AND(D6>=12,D6<=24),">12",IF(D6>24,">24","<12"))'  # Example formula (modify as needed)

    # Get the last row of data in column A
    last_row = pivot_sheet.range('D' + str(pivot_sheet.cells.last_cell.row)).end('up').row

    # Define the range to autofill
    autofill_range = pivot_sheet.range(f'E6:E{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #EXEMPT
    formula_cell = pivot_sheet.range('F6')  # Cell where the formula starts
    formula_cell.value = '=IF(C6>=4,"Yes",IF(B6>6,"Yes","No"))'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'F6:F{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #F.O. Cable
    formula_cell = pivot_sheet.range('G6')  # Cell where the formula starts
    formula_cell.value = '=COUNTIFS(NTR60!$D:$D,NTR_60!$A6,NTR60!$AC:$AC,NTR_60!G$5)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'G6:G{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #Coax Cable
    formula_cell = pivot_sheet.range('H6')  # Cell where the formula starts
    formula_cell.value = '=COUNTIFS(NTR60!$D:$D,NTR_60!$A6,NTR60!$AC:$AC,NTR_60!H$5)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'H6:H{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #Electronic
    formula_cell = pivot_sheet.range('I6')  # Cell where the formula starts
    formula_cell.value = '=COUNTIFS(NTR60!$D:$D,NTR_60!$A6,NTR60!$AC:$AC,NTR_60!I$5)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'I6:I{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #Power
    formula_cell = pivot_sheet.range('J6')  # Cell where the formula starts
    formula_cell.value = '=COUNTIFS(NTR60!$D:$D,NTR_60!$A6,NTR60!$AC:$AC,NTR_60!J$5)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'J6:J{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #Blanks
    formula_cell = pivot_sheet.range('K6')  # Cell where the formula starts
    formula_cell.value = '=COUNTIFS(NTR60!$D:$D,NTR_60!$A6,NTR60!$AC:$AC,NTR_60!K$5)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'K6:K{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    #Total
    formula_cell = pivot_sheet.range('L6')  # Cell where the formula starts
    formula_cell.value = '=SUM(G6:J6)'  # Example formula (modify as needed)
    autofill_range = pivot_sheet.range(f'L6:L{last_row}')
    formula_cell.api.AutoFill(autofill_range.api, xw.constants.AutoFillType.xlFillDefault)

    app.calculation = 'automatic'
    app.screen_updating = True


    print("DONE APPLYING FORMULAS.")

    # Save WB with the updated data
    WB.save(file_path)


    

# =====================================
# SECTION: NTR INCOMPLETE DETAILS
# =====================================

print("QUERYING NTR INCOMPLETE. . . . . . . . . .")
   
# SQL command to read the data
sqlQuery = """
SELECT *
FROM (
SELECT
    n.NTR_ID AS 'Ref#',
    CONCAT(DATE_FORMAT(n.NTR_DATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TIME, '%H:%i')) AS 'Downtime',
    n.NTR_NODE AS 'Node',
    s.NTR_NODE AS 'Per Node',
    n.NTR_AMPLIFIER AS 'Amplifier',
    n.NTR_LOC AS 'Location',
    n.NTR_SYSTEM AS 'System',
    n.NTR_SMS AS 'SMS',
    n.NTR_PTCODE AS 'PT Code',
    n.NTR_SERVICE AS 'Service Affected',
    n.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    CONCAT(DATE_FORMAT(n.NTR_TRDATESTART, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TRTIMESTART, '%H:%i')) AS 'TS Start',
    s.TS_END AS 'TS End',
    n.NTR_ETR AS 'ETR',
    ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) AS 'Duration',
    n.NTR_TROUBLEDETAILS AS 'Trouble Details',
    n.NTR_RESTOREDETAILS AS 'Restore Details',
    n.NTR_EXCLUDE AS 'Exclude',
    n.NTR_EXCLUDE_CAT AS 'Exclusion Reason',
    n.NTR_FAULTNUM AS 'NF#',
    n.NTR_FAULTSTS AS 'NF Status',
    CONCAT(DATE_FORMAT(n.NTR_CRDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_CRTIME, '%H:%i')) AS 'Creation Date/Time',
    CONCAT(DATE_FORMAT(n.NTR_1CDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_1CTIME, '%H:%i')) AS '1st Call Date/Time',
    n.NTR_GROUP AS 'Group',
    n.NTR_OIC AS 'OIC',
    n.NTR_TEAM AS 'Team',
    n.NTR_PORIGIN AS 'Point of Origin',
    n.NTR_CCD_REASON AS 'CCD Reason',
    n.NTR_CATEGORY AS 'Category',
    n.NTR_CAUSE AS 'Cause',
    n.NTR_CONTROLLABILITY AS 'Controllability',
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'CATV' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected CATV`,
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'SBB' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected SBB`,
        (SELECT COUNT(*) 
        FROM device_db.device_tbl2 
        WHERE node_id = n.NTR_NODE 
        AND status_node = '1' 
        AND cdn = '1W') AS 'Device Affected CATV',
        (SELECT COUNT(*) 
        FROM device_db.device_tbl2 
        WHERE node_id = n.NTR_NODE 
        AND status_node = '1' 
        AND cdn = '2W') AS 'Device Affected SBB',
    n.NTR_ONBOARD AS 'Detected by',
    n.NTR_RELAYTIME AS 'Relaytime',
    n.NTR_CORPO AS 'Is_corpo',
    n.NTR_CORPO_LIST AS 'Corpo List',
    n.NTR_GROUP AS 'Area',
    n.NTR_ALARM_TRIGGER AS 'Alarm Trigger',
    n.NTR_ALARM_SOURCE AS 'Alarm Source',
    n.NTR_TIMESTAMP AS 'Detection Time'
FROM noc_db.ntr_tbl n
JOIN noc_db.ntr_subs_affected_tbl s ON n.NTR_ID = s.NTR_ID
WHERE TIMESTAMP(n.NTR_DATE, n.NTR_TIME) >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY)
  AND TIMESTAMP(n.NTR_DATE, n.NTR_TIME) < CURRENT_DATE
  AND s.NTR_NODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
  AND s.ACTIVE = 1
  AND n.NTR_STATUS = 'OK'
  AND (s.TS_END IS NULL OR TRIM(s.TS_END) = '')    -- No TS End
GROUP BY n.NTR_ID, n.NTR_NODE, s.NTR_NODE, n.NTR_AMPLIFIER, n.NTR_LOC, n.NTR_SYSTEM, 
         n.NTR_SMS, n.NTR_PTCODE, n.NTR_SERVICE, n.NTR_STATUS, n.NTR_TRDATESTART, 
         n.NTR_TRTIMESTART, s.TS_END, n.NTR_ETR, n.NTR_TROUBLEDETAILS, n.NTR_RESTOREDETAILS, 
         n.NTR_EXCLUDE, n.NTR_EXCLUDE_CAT, n.NTR_FAULTNUM, n.NTR_FAULTSTS, n.NTR_CRDATE, 
         n.NTR_CRTIME, n.NTR_1CDATE, n.NTR_1CTIME, n.NTR_GROUP, n.NTR_OIC, n.NTR_TEAM, 
         n.NTR_PORIGIN, n.NTR_CCD_REASON, n.NTR_CATEGORY, n.NTR_CAUSE, n.NTR_CONTROLLABILITY, 
         n.NTR_ONBOARD, n.NTR_RELAYTIME, n.NTR_CORPO, n.NTR_CORPO_LIST, n.NTR_ALARM_TRIGGER, 
         n.NTR_ALARM_SOURCE, n.NTR_TIMESTAMP

UNION

SELECT
    n.NTR_ID AS 'Ref#',
    CONCAT(DATE_FORMAT(n.NTR_DATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TIME, '%H:%i')) AS 'Downtime',
    n.NTR_NODE AS 'Node',
    s.NTR_NODE AS 'Per Node',
    n.NTR_AMPLIFIER AS 'Amplifier',
    n.NTR_LOC AS 'Location',
    n.NTR_SYSTEM AS 'System',
    n.NTR_SMS AS 'SMS',
    n.NTR_PTCODE AS 'PT Code',
    n.NTR_SERVICE AS 'Service Affected',
    n.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    CONCAT(DATE_FORMAT(n.NTR_TRDATESTART, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TRTIMESTART, '%H:%i')) AS 'TS Start',
    s.TS_END AS 'TS End',
    n.NTR_ETR AS 'ETR',
    ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) AS 'Duration',
    n.NTR_TROUBLEDETAILS AS 'Trouble Details',
    n.NTR_RESTOREDETAILS AS 'Restore Details',
    n.NTR_EXCLUDE AS 'Exclude',
    n.NTR_EXCLUDE_CAT AS 'Exclusion Reason',
    n.NTR_FAULTNUM AS 'NF#',
    n.NTR_FAULTSTS AS 'NF Status',
    CONCAT(DATE_FORMAT(n.NTR_CRDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_CRTIME, '%H:%i')) AS 'Creation Date/Time',
    CONCAT(DATE_FORMAT(n.NTR_1CDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_1CTIME, '%H:%i')) AS '1st Call Date/Time',
    n.NTR_GROUP AS 'Group',
    n.NTR_OIC AS 'OIC',
    n.NTR_TEAM AS 'Team',
    n.NTR_PORIGIN AS 'Point of Origin',
    n.NTR_CCD_REASON AS 'CCD Reason',
    n.NTR_CATEGORY AS 'Category',
    n.NTR_CAUSE AS 'Cause',
    n.NTR_CONTROLLABILITY AS 'Controllability',
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'CATV' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected CATV`,
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'SBB' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected SBB`,
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '1W') AS 'Device Affected CATV',
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '2W') AS 'Device Affected SBB',
    n.NTR_ONBOARD AS 'Detected by',
    n.NTR_RELAYTIME AS 'Relaytime',
    n.NTR_CORPO AS 'Is_corpo',
    n.NTR_CORPO_LIST AS 'Corpo List',
    n.NTR_GROUP AS 'Area',
    n.NTR_ALARM_TRIGGER AS 'Alarm Trigger',
    n.NTR_ALARM_SOURCE AS 'Alarm Source',
    n.NTR_TIMESTAMP AS 'Detection Time'
FROM noc_db.ntr_tbl n
JOIN noc_db.ntr_subs_affected_tbl s ON n.NTR_ID = s.NTR_ID
WHERE 
    n.NTR_DATE >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY) 
    AND n.NTR_DATE < CURRENT_DATE
    AND s.NTR_NODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
    AND s.ACTIVE = 1
    AND n.NTR_STATUS = 'OK'
    AND ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) < 0    -- Negative Duration
GROUP BY n.NTR_ID, n.NTR_NODE, s.NTR_NODE, n.NTR_AMPLIFIER, n.NTR_LOC, n.NTR_SYSTEM, 
         n.NTR_SMS, n.NTR_PTCODE, n.NTR_SERVICE, n.NTR_STATUS, n.NTR_TRDATESTART, 
         n.NTR_TRTIMESTART, s.TS_END, n.NTR_ETR, n.NTR_TROUBLEDETAILS, n.NTR_RESTOREDETAILS, 
         n.NTR_EXCLUDE, n.NTR_EXCLUDE_CAT, n.NTR_FAULTNUM, n.NTR_FAULTSTS, n.NTR_CRDATE, 
         n.NTR_CRTIME, n.NTR_1CDATE, n.NTR_1CTIME, n.NTR_GROUP, n.NTR_OIC, n.NTR_TEAM, 
         n.NTR_PORIGIN, n.NTR_CCD_REASON, n.NTR_CATEGORY, n.NTR_CAUSE, n.NTR_CONTROLLABILITY, 
         n.NTR_ONBOARD, n.NTR_RELAYTIME, n.NTR_CORPO, n.NTR_CORPO_LIST, n.NTR_ALARM_TRIGGER, 
         n.NTR_ALARM_SOURCE, n.NTR_TIMESTAMP

UNION

SELECT
    n.NTR_ID AS 'Ref#',
    CONCAT(DATE_FORMAT(n.NTR_DATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TIME, '%H:%i')) AS 'Downtime',
    n.NTR_NODE AS 'Node',
    s.NTR_NODE AS 'Per Node',
    n.NTR_AMPLIFIER AS 'Amplifier',
    n.NTR_LOC AS 'Location',
    n.NTR_SYSTEM AS 'System',
    n.NTR_SMS AS 'SMS',
    n.NTR_PTCODE AS 'PT Code',
    n.NTR_SERVICE AS 'Service Affected',
    n.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    CONCAT(DATE_FORMAT(n.NTR_TRDATESTART, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TRTIMESTART, '%H:%i')) AS 'TS Start',
    s.TS_END AS 'TS End',
    n.NTR_ETR AS 'ETR',
    ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) AS 'Duration',
    n.NTR_TROUBLEDETAILS AS 'Trouble Details',
    n.NTR_RESTOREDETAILS AS 'Restore Details',
    n.NTR_EXCLUDE AS 'Exclude',
    n.NTR_EXCLUDE_CAT AS 'Exclusion Reason',
    n.NTR_FAULTNUM AS 'NF#',
    n.NTR_FAULTSTS AS 'NF Status',
    CONCAT(DATE_FORMAT(n.NTR_CRDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_CRTIME, '%H:%i')) AS 'Creation Date/Time',
    CONCAT(DATE_FORMAT(n.NTR_1CDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_1CTIME, '%H:%i')) AS '1st Call Date/Time',
    n.NTR_GROUP AS 'Group',
    n.NTR_OIC AS 'OIC',
    n.NTR_TEAM AS 'Team',
    n.NTR_PORIGIN AS 'Point of Origin',
    n.NTR_CCD_REASON AS 'CCD Reason',
    n.NTR_CATEGORY AS 'Category',
    n.NTR_CAUSE AS 'Cause',
    n.NTR_CONTROLLABILITY AS 'Controllability',
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'CATV' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected CATV`,
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'SBB' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected SBB`,
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '1W') AS 'Device Affected CATV',
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '2W') AS 'Device Affected SBB',
    n.NTR_ONBOARD AS 'Detected by',
    n.NTR_RELAYTIME AS 'Relaytime',
    n.NTR_CORPO AS 'Is_corpo',
    n.NTR_CORPO_LIST AS 'Corpo List',
    n.NTR_GROUP AS 'Area',
    n.NTR_ALARM_TRIGGER AS 'Alarm Trigger',
    n.NTR_ALARM_SOURCE AS 'Alarm Source',
    n.NTR_TIMESTAMP AS 'Detection Time'
FROM noc_db.ntr_tbl n
JOIN noc_db.ntr_subs_affected_tbl s ON n.NTR_ID = s.NTR_ID
WHERE 
    n.NTR_DATE >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY) 
    AND n.NTR_DATE < CURRENT_DATE
    AND s.NTR_NODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
    AND s.ACTIVE = 1
    AND n.NTR_STATUS = 'OK'
    AND n.NTR_CATEGORY IN ('NULL', '-', ' ')   -- No Category
GROUP BY n.NTR_ID, n.NTR_NODE, s.NTR_NODE, n.NTR_AMPLIFIER, n.NTR_LOC, n.NTR_SYSTEM, 
         n.NTR_SMS, n.NTR_PTCODE, n.NTR_SERVICE, n.NTR_STATUS, n.NTR_TRDATESTART, 
         n.NTR_TRTIMESTART, s.TS_END, n.NTR_ETR, n.NTR_TROUBLEDETAILS, n.NTR_RESTOREDETAILS, 
         n.NTR_EXCLUDE, n.NTR_EXCLUDE_CAT, n.NTR_FAULTNUM, n.NTR_FAULTSTS, n.NTR_CRDATE, 
         n.NTR_CRTIME, n.NTR_1CDATE, n.NTR_1CTIME, n.NTR_GROUP, n.NTR_OIC, n.NTR_TEAM, 
         n.NTR_PORIGIN, n.NTR_CCD_REASON, n.NTR_CATEGORY, n.NTR_CAUSE, n.NTR_CONTROLLABILITY, 
         n.NTR_ONBOARD, n.NTR_RELAYTIME, n.NTR_CORPO, n.NTR_CORPO_LIST, n.NTR_ALARM_TRIGGER, 
         n.NTR_ALARM_SOURCE, n.NTR_TIMESTAMP

UNION

SELECT
    n.NTR_ID AS 'Ref#',
    CONCAT(DATE_FORMAT(n.NTR_DATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TIME, '%H:%i')) AS 'Downtime',
    n.NTR_NODE AS 'Node',
    s.NTR_NODE AS 'Per Node',
    n.NTR_AMPLIFIER AS 'Amplifier',
    n.NTR_LOC AS 'Location',
    n.NTR_SYSTEM AS 'System',
    n.NTR_SMS AS 'SMS',
    n.NTR_PTCODE AS 'PT Code',
    n.NTR_SERVICE AS 'Service Affected',
    n.NTR_STATUS AS 'Status', -- Joined column from ntr_tbl
    CONCAT(DATE_FORMAT(n.NTR_TRDATESTART, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_TRTIMESTART, '%H:%i')) AS 'TS Start',
    s.TS_END AS 'TS End',
    n.NTR_ETR AS 'ETR',
    ROUND(TIMESTAMPDIFF(SECOND, 
                        TIMESTAMP(n.NTR_DATE, n.NTR_TIME), 
                        TIMESTAMP(s.TS_END, '00:00:00')) / 3600, 2) AS 'Duration',
    n.NTR_TROUBLEDETAILS AS 'Trouble Details',
    n.NTR_RESTOREDETAILS AS 'Restore Details',
    n.NTR_EXCLUDE AS 'Exclude',
    n.NTR_EXCLUDE_CAT AS 'Exclusion Reason',
    n.NTR_FAULTNUM AS 'NF#',
    n.NTR_FAULTSTS AS 'NF Status',
    CONCAT(DATE_FORMAT(n.NTR_CRDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_CRTIME, '%H:%i')) AS 'Creation Date/Time',
    CONCAT(DATE_FORMAT(n.NTR_1CDATE, '%d-%b-%y'), ' ', DATE_FORMAT(n.NTR_1CTIME, '%H:%i')) AS '1st Call Date/Time',
    n.NTR_GROUP AS 'Group',
    n.NTR_OIC AS 'OIC',
    n.NTR_TEAM AS 'Team',
    n.NTR_PORIGIN AS 'Point of Origin',
    n.NTR_CCD_REASON AS 'CCD Reason',
    n.NTR_CATEGORY AS 'Category',
    n.NTR_CAUSE AS 'Cause',
    n.NTR_CONTROLLABILITY AS 'Controllability',
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'CATV' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected CATV`,
    SUM(CASE WHEN s.SERVICE_AFFECTED = 'SBB' THEN s.SUBS_COUNT ELSE 0 END) AS `Subs Affected SBB`,
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '1W') AS 'Device Affected CATV',
    (SELECT COUNT(*) 
     FROM device_db.device_tbl2 
     WHERE node_id = n.NTR_NODE 
     AND status_node = '1' 
     AND cdn = '2W') AS 'Device Affected SBB',
    n.NTR_ONBOARD AS 'Detected by',
    n.NTR_RELAYTIME AS 'Relaytime',
    n.NTR_CORPO AS 'Is_corpo',
    n.NTR_CORPO_LIST AS 'Corpo List',
    n.NTR_GROUP AS 'Area',
    n.NTR_ALARM_TRIGGER AS 'Alarm Trigger',
    n.NTR_ALARM_SOURCE AS 'Alarm Source',
    n.NTR_TIMESTAMP AS 'Detection Time'
FROM noc_db.ntr_tbl n
JOIN noc_db.ntr_subs_affected_tbl s ON n.NTR_ID = s.NTR_ID
WHERE 
    n.NTR_DATE >= DATE_SUB(CURRENT_DATE, INTERVAL 60 DAY) 
    AND n.NTR_DATE < CURRENT_DATE
    AND s.NTR_NODE NOT IN ('BIZ', 'DITO', 'DITOTEL')
    AND s.ACTIVE = 1
    AND n.NTR_STATUS = 'OK'
    AND n.NTR_CAUSE IN ('NULL', '-', ' ')     -- No Cause
GROUP BY n.NTR_ID, n.NTR_NODE, s.NTR_NODE, n.NTR_AMPLIFIER, n.NTR_LOC, n.NTR_SYSTEM, 
         n.NTR_SMS, n.NTR_PTCODE, n.NTR_SERVICE, n.NTR_STATUS, n.NTR_TRDATESTART, 
         n.NTR_TRTIMESTART, s.TS_END, n.NTR_ETR, n.NTR_TROUBLEDETAILS, n.NTR_RESTOREDETAILS, 
         n.NTR_EXCLUDE, n.NTR_EXCLUDE_CAT, n.NTR_FAULTNUM, n.NTR_FAULTSTS, n.NTR_CRDATE, 
         n.NTR_CRTIME, n.NTR_1CDATE, n.NTR_1CTIME, n.NTR_GROUP, n.NTR_OIC, n.NTR_TEAM, 
         n.NTR_PORIGIN, n.NTR_CCD_REASON, n.NTR_CATEGORY, n.NTR_CAUSE, n.NTR_CONTROLLABILITY, 
         n.NTR_ONBOARD, n.NTR_RELAYTIME, n.NTR_CORPO, n.NTR_CORPO_LIST, n.NTR_ALARM_TRIGGER, 
         n.NTR_ALARM_SOURCE, n.NTR_TIMESTAMP
) AS combined_result
ORDER BY `Downtime`;
 
"""

# Getting the data from SQL into a pandas DataFrame
connection = pyodbc.connect(**db_config)
df = pd.read_sql(sql=sqlQuery, con=connection)
 

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







