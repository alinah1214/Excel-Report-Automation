# Rectifier Setting Issue Report Automation using VBA

### Project Overview
This project demonstrates automated reporting in Excel by utilizing real data from Telenor cellular network sites experiencing rectifier setting problems. The report includes sites with rectifier setting issue alarms lasting longer than 3 days that haven't been resolved yet.

### Data Sources
In order to make this report, the data is taken from two different Network Management Systems.
1. EMS
2. Netact
- Three files are from EMS (Ems1.csv, EMS2.csv, EMS3.csv)
- Three files are from Netact (Netact1.csv, Netact2.csv)

### Report Steps
The macro takes the following steps in order to make the report.

1. The data from the three EMS files is aggregated together, and all columns except the required ones—Severity, Raised time, NE, and Alarm Code—are removed
2. The data is cleaned by filtering out DTP, TOF, LCK, and NULL sites from the NE column.
3. The six digits are extracted from NE column, and a separate column is created to represent the site ID.
4. Within the macro sheet, there is a region column corresponding to the site ID. From this sheet, the region is looked up using VLOOKUP in the agrregated EMS data.
6. The aging of the alarms is calculated in terms of days in a separate column, and then the sites having an age of less than 4 days are removed from the report.
7. A Pivot table is created containing the count of rectifier setting issue alarms region-wise
8. The resulting report contains the data of sites with rectifier setting issues and a Pivot table is generated under the name **ZTE Land**.

#### Note: 
The same steps are taken for the Netact sheets, and in the end, the report for Netact data is generated under the name **Nokia Land**.


### Macro Usage
- Place EMS1, EMS2, EMS3, Netact1, Netact2, and Rectifier Setting Issue Macro in one folder.
- Open the Rectifier Setting Issue macro and click on the **Run** button.
- The report will be generated within 2-3 seconds.


### Output Files:
Two files are generated in output as a result of running the macro.
1. ZTE Land.xlsx (For EMS)
2. Nokia Land.xlsx (For Netact)
