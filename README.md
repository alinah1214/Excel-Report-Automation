# Rectifier Setting Issue Report Automation using VBA

### Project OverView
This project basically domonstrates the report automation in Excel by taking the real data of telenor cellular network sites having the rectifier setting problem. The report actually have all the sites that are having the recifier setting issue alarm longer than 3 days and is not resloved yet. 

### Data Sources
 The data is collected from two different network management systems.
1. EMS
2. Netact
Three files are from EMS (Ems1, EMS2, EMS3)
Three files are from Netact (Netact1, Netac2)

Output Files:
Two files are generated in output as a reult of running a macro.
1. ZTE Land (For EMS)
2. Nokia Land (For Netact)

### Report Making Steps 
The following are taken in order to makes the report.

1. The Three EMS files data uis aggregated together
2. The data is cleaned by removing the DTP, TOF, LCK and NULL sites from the NE coloumn by using filter
3. The aging of the alarms is calculated in terms of days in separarte column and then 
4. 
