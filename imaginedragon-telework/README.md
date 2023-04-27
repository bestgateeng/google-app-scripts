# IMAGINEDRAGON Telework Log Emailer

## Usage
This project supports the emailing of the IMAGINEDRAGON program's telework logs. The script expects a Google Drive Folder that contains 1:M Google Sheets files. Every file is added as an attachment. The Google App Script in whcih this runs is configured to execute on the 1st of every month between 0900-1000 EST.

## Install notes
1. This simply runs as a Google App Script.
2. A Trigger is currently configured to execute on the 1st of each month between 0900-1000 EST.
3. The `FOLDER_ID` value is not in this source-controlled version, and should be added into the production App Script.

## Tips for tracking time
A formula that can be applied to Column E of the sheets to convert start/stop times in military time into decimal: 
`=MROUND(DOLLARDE(C2/100,60)-DOLLARDE(B2/100,60),0.25)`

A formula that can be applied to any cell to the right of Column E, which creates a table of total hours worked for each day (expects the above formula to be applied in Column E):
`=QUERY(A:E, "SELECT A, SUM(E) WHERE A IS NOT NULL GROUP BY A LABEL A 'Unique Date', SUM(E) 'Hours Worked'", 0)`
