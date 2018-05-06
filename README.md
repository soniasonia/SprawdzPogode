# SprawdzPogode
Project created in purpose of learning C# and OOP (SOLID & DRY way):
- Basic data input features 
- Excel handling
- Selenium.

In this training program you read directories from XML file: 
- input TXT file
- output Excel file.

In input TXT file you have a list of cities for which you search Google (city name + weather):
```
  Katowice
  Amsterdam
  Praga
```
Every search is saved to Excel file as a new row. If the weather DIV is visible on the top of results list for your search, you extract current temperature, precipitation and wind and save them to Excel. 
The columns in Excel file are:
- Time (yyyy/MM/dd HH:mm:ss)
- City 
- Temperature
- Precipitation 
- Wind
- Status (Success of Fail)
