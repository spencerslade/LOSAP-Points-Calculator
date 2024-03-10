# LOSAP Points Calculator (LPC)

**What is LOSAP?**

The General Municipal New York Laws, Article 11-A, defines Contribution Service Award Programs for Volunteer Ambulance Workers. Read more [here](https://www.lawserver.com/law/state/new-york/ny-laws/new_york_laws_general_municipal_article_11-aa). Many Volunteer Ambulance Corps have accordingly implemented a Length of Service Awards Program (LOSAP).

**The purpose of the 'LOSAP Points Calculator'**

This program implements quick and easy-to-use method to aggregate Ambulance Worker data from multiple sources, and to calculate points based on a predefined scheme. Specifically, data are drawn from signup databases as well as self-reported data from a collection of custom Excel spreadsheets.

# How to use the program

Using the program should be quite straightforward. Essentially, three different sources of data are imported individually (in any order). A data structure is produced that combines these data. The summary data can be saved as an Excel file.

## Import the signup hours from ‘[I am Responding](https://www.iamresponding.com/)’ records

An export file is generated in Excel format for a chosen time period (i.e. for the month of January). A sum total of signup hours for each person is determined, and "**Tour of Duty" points** are calculated accordingly (1 point for each hour).

Note that the structure of the Excel spreadsheet to be read should be taken in consideration. For example, the headings may be present in row 3, in which case the first two rows should be skipped. The raw data are valid until a certain row (e.g. row 251), before an aggregated form of the data are repeated. In this case, we want to **skip 2 rows and read until row 251**. These are the default numbers, but can be adjusted in the Settings dialog (Edit -\> Settings).

## Import the number of calls responded to from [Electronic Patient Care Reporting](https://www.ems1.com/ems-products/ePCR-Electronic-Patient-Care-Reporting/) (ePCR) records

An export file is generated in CSV (comma separated values) format from ePCR data for a given time period (e.g. for the month of January). The number of calls per person is calculated from these data to calculate the "Calls Responded To" points (0.5 points to each call responded to).

## Import member-submitted Excel spreadsheets (self-reported data)

An Excel spreadsheet was made that members can use to self-report their data. The Excel spreadsheet has a specific layout and contains the following information:

|             | Excel cell |
|-------------|------------|
| Member name | D4         |
| Duty hours  | E7         |
| Total calls | E8         |

Specific activities are listed from row 11 onwards (i.e. the first 10 rows are skipped when reading these data). Activities are categorized as “Training Course”, “Drills, CMEs”, “Meetings”, “Miscellaneous”, and “Disability”. The Excel sheet name to be read is “point tracker”. The 'LOSAP Points Calculator' software assumes these values, but they can be changed in the Settings dialog (Edit -\> Settings).

The 'LOSAP Points Calculator' assumes that **all spreadsheets for a given period (e.g. for the month of January) are all be present in the same folder**. The 'LOSAP Points Calculator' will open each Excel spreadsheet, read all data from each spreadsheet, group data as needed and calculate points based on reported hours using a predefined formula.

## Export the results to Excel file

After all data have been imported from the various sources, the aggregated points can be saved to an Excel file. This file is in the required format for reporting and auditing purposes. Two additional columns are added, named “SR_Signup” and “SR_Calls”, which respectively report points based on the self-reported member-submitted Excel spreadsheets.

Note that no checks are performed whether data from all three sources have been imported.

## Other functions

Data can be cleared and the program reset to its startup conditions by “File -\> New” or “Edit -\> Clear”

# License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see \<http://www.gnu.org/licenses/\>.
