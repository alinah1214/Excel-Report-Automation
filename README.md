# Rectifier Setting Issue Report Automation using VBA

### Project OverView
This project basically domonstrates the report automation in Excel by taking the real data of telenor cellular network sites having the rectifier setting problem. The report actually have all the sites that are having the recifier setting issue alarm longer than 3 days and is not resloved yet. 

### Data Sources
 The data is collected from two different network management systems.
1. EMS
2. Netact
Three files are from EMS (Ems1, EMS2, EMS3)
Three files are from Netact (Netact1, Netact2)

### Report Steps
The macro takes the following steps in order to makes the report.

1. The Three EMS files data is aggregated together and all columns are removed except the required coloumns which are Seveirty, Raised time, NE and Alram Code.
2. The data is cleaned by removing the DTP, TOF, LCK and NULL sites from the NE coloumn by using filteration
3. The 6 digits are extracted and a seprate coloumn is made which actually represents sites ID.
4. within macro sheet, there is region coloumn against sites ID so from this sheet the region is vlookuped in EMS sheet.
6. The aging of the alarms is calculated in terms of days in separarte column and then the sites haiving aging < 4 days are removed from the report.
7. The Pivit table is made containing the count of rectifier setting issue alarms region wise
8. The dats is Named in a sheet named **ZTE land**.

#### Note: 
Same steps are taken for Netact sheets and in the end the report for Netact data is generted by the name **Nokia Land**.


### Macro Usage
Place EMS1, EMS2, EMS3, Naetact1, Netact2 and Rectifier setting Issue Macro in one folder
Open Rectifier Setting Isuue macro and click on **Run** button
The reoprt will be gereneted within 2-3 senonds


### Output Files:
Two files are generated in output as a reult of running a macro.
1. ZTE Land (For EMS)
2. Nokia Land (For Netact)
