# High Level Testing

| **Test ID** | **Description**                                              | **Exp IN** | **Exp OUT** | **Actual Out** |**Type Of Test**  |    
|-------------|--------------------------------------------------------------|------------|-------------|----------------|------------------|
| 1 | To Access and read the workbook stored in same folder/location | Excel file with data | Read the excel file Number of worksheets=1, Number of sheets=5, Number of Rows and Columns=40x10 Excel file = Data1.xlsx| Excel file with all its sheet is accessible. | Initial Testing |
| 2 | Master Sheet Creation | User input: Name, PS No. and E-mail ID. | Master Sheet creation in the existing Excel file. | Master Sheet created in existing Excel file. |Requirement based |
| 3 | To Search by Name, PS no. and E-mail ID. | User input: Name, PS No. and E-mail ID | If Data matched with User Input: Copy and Paste all the related data in all sheets into “Master Sheet” If Data not matched with User Input: Print No such data Present in Data Base and not append any data to “Master Sheet”. | All data copied and pasted to Master Sheet if User Input matched else no data copy. | Requirement based |


# Low Level Testing

| **Test ID** | **Description**                                              | **Exp IN** | **Exp OUT** | **Actual Out** |**Type Of Test**  |    
|-------------|--------------------------------------------------------------|------------|-------------|----------------|------------------|
| 1 |	To access the worksheet by providing the path. |	File Path: ‘Data1.xlsx’ Keep all files in same location. Note: Already provided in the code. |	Workbook (.xlsx) load without error. | Workbook loaded without errors |	Initial |
| 2 |	To access data of all sheets in the Workbook. |	File Path: ‘Data1.xlsx’ Keep all files in same location. Note: Already provided in the code |	Starts reading every sheet from the worksheet |	Reading all sheets | Scenario based |
| 3	| Searching data by Name, PS No. and Email ID. (all correct details) | Enter the name for Data 1: Jordan Cassey Enter the PS No for Data 1: 99003760 Enter email id for Data 1: j.casey@ltts.com |	Print “Data Present in DataBase”. Data from all sheets matching with user input send to Master Sheet. |	“Data Present in DataBase” printed. Data from all sheets matching with user input is send to Master Sheet. | Requirement based |
| 4 | Searching data by Name, PS No. and Email ID. (Incorrect Name) |	Enter the name for Data 1: abcdef Enter the PS No for Data 1: 99003760 Enter email id for Data 1: j.casey@ltts.com | Print “Data Provided NOT FOUND in DataBase”. No data pasted to Master Sheet. |	“Data Provided NOT FOUND in DataBase” printed. No data pasted to Master Sheet. |	Requirement based |
