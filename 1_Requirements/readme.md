# Requirements
## Introduction
It is a Data Set problem that will allow to retrive data from data set. However, the input has form of the PS No.,name and email id. And output is  all student data .

# Detail requirements
## High Level Requirements:
|id  |Requirements  | Description  |Status  |
| --- | --- | --- | --- |
|HL1 | Search data from sheet |Search all data from sheets when user gives the name, PS No. and email id to be searched.|IMPLEMENTED|
|HL2 | write data into new Sheet  | Write all the data from different sheets in one Master Sheet|IMPLEMENTED |
|HL3 |Extract data from sheets using given input|Write new required data in the excel file. |IMPLEMENTED |



##  Low level Requirements:

|id  |Requirments  | Description  |Status  |
| --- | --- | --- | --- |
|LL001 | Data Collection |worksheets contains the data of company details and academic details of users input|IMPLEMENTED
|LL002 | Each Sheet Contains 10 Column and 40 Rows |Each Sheet showing 10X40 format|IMPLEMENTED |
|LL003 | Excel file format | the workbook file should be of .xslx format|IMPLEMENTED
|LL004 |Inputs|User can give multiple inputs like name,PS No, name amd email id at once|IMPLEMENTED
|LL005 |Reading Data|Reading all 5 worksheets from workbook|IMPLEMENTED
|LL004 |Searching Data|Search for specific data based on user specific inputs|IMPLEMENTED
|LL006 | Master Sheet Contains Created  | Master Sheet Contains 40X40 Format|IMPLEMENTED |

  
## SWOT ANALYSIS

![SWOT Analysis](https://user-images.githubusercontent.com/78858575/111780469-78287780-88dd-11eb-8438-2637230c6579.png)
 
# 4W1H

## Why:
* We are using to retrieve the data of an individual candidate from the excel workbook of 5 sheets where all the relevant data of 40 candidates is present.
* We can easily access the details of that individual by giving some input such as name, Ps no and email id.



## What:
 *	We are preparing the master excel sheet to search and retrieve data.
 *	For easy search of a particular cell or data of a person.
 *	Provides information of every person details like bio,academics,health.
 

## When:
*	Evalaution of exams.
*	Searching for person information.
*	Contact information.



## Where:
*	To check the information and bio of a person.
*	Very useful during emergency times.
*	We can also use it for evaluation of marks using the mail,search the location of person.



## How:
*	Input:- giving the ps no,name or email of person
*	Output: -gives all the relevant information of person related to ps no, name or email
*	source: -excel sheets  give the output into the master sheet
