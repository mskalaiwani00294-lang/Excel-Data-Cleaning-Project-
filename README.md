# Project Title

Data Cleaning and Preprocessing Using Microsoft 
Excel

## Description

This project focuses on cleaning and preprocessing a product dataset using Microsoft Excel. The objective is to improve data quality by handling missing values, correcting inconsistent text formats, fixing category typos, removing duplicate records, splitting and merging columns, applying number formatting, and using conditional formatting for better visualization. The cleaned dataset ensures accuracy, consistency, and reliability for further analysis and reporting.

## Getting Started

### Dependencies
Before starting the project, ensure the following requirements are available: • 
Microsoft Excel 2016 or later • Windows 10/11 or macOS • Basic knowledge of Excel 
formulas and spreadsheet operations


### Installing

Steps to Set Up the Project 
1. Download or create the Excel dataset file. 
2. Open the dataset using Microsoft Excel. 
3. Ensure all columns are properly labeled: o Product ID o Product Name o Brand 
Name o Price o Quantity o Category 
4. Save the file in .xlsx format. No additional software installation is required 

### Executing program

Open the dataset in Microsoft Excel.
Identify and handle missing values in the Price and Category columns.
Use Find and Replace to correct text inconsistencies and category typos.
Remove duplicate rows using the Remove Duplicates feature.
Split the Product ID column into Manufacturing Date and Country Code.
Merge Brand Name and Product Name into Product Brand.
Apply currency formatting to the Price column.
Format Manufacturing Date as DD-MM-YYYY.
Apply conditional formatting to the Price and Category columns.
Save the cleaned dataset.
Example formulas used:
Excel Formula
=AVERAGE(E:E)
=LEFT(A2,9)
=RIGHT(A2,2)
=B2&" "&C2
=PROPER(D2)
Help
Common Issues
Problem: Dates are displayed incorrectly.
Solution: Format the column using:
Plain text
DD-MM-YYYY
Problem: Text formatting remains inconsistent.
Solution: Use the PROPER function:
Excel Formula
=PROPER(A2)

## Help

Duplicate records still exist.
Solution: Select the entire dataset before using Remove Duplicates


## Authors

Contributors names and contact info
M.Kalaivani 
Data  Analytics  Project 

ex. Dominique Pizzie  
ex. [@DomPizzie](https://twitter.com/dompizzie)

## Version History

* 0.2
    Initial project completion
Missing values handled
Text inconsistencies corrected
Duplicate records removed
Columns split and merged
Number formatting applied
Conditional formatting implemented

* 0.1
    * Initial Release

## License

This project is licensed under the [M.Kalaivani] License - see the LICENSE.md file for details

## Acknowledgments

Inspiration, code snippets, etc.
* [awesome-readme](https://github.com/matiassingers/awesome-readme)
* [PurpleBooth](httMicrosoft Excel Documentation
Data Analytics Training Program
Excel Functions and Data Cleaning Tutorials
Course instructors and mentors for project guidance
Save this content as README.md inside your project folder along with:
Cleaned Excel file (.xlsx)
PDF report with screenshots and explanations
ps://gist.github.com/PurpleBooth/109311bb0361f32d87a2)
* [dbader](https://github.com/dbader/readme-template)
* [zenorocha](https://gist.github.com/zenorocha/4526327)
* [fvcproductions](https://gist.github.com/fvcproductions/1bfc2d4aecb01a834b46)
