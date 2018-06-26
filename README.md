# SprawdzPogode
Project created in purpose of learning C# and OOP (SOLID & DRY way):
- Basic data input features 
- Excel handling
- Selenium.

### How does it work
Read directories from XML config file: 
- input TXT file
- output Excel file.

In input TXT file you have a list of cities for which you search Google (city name + weather):
```
  Katowice
  Amsterdam
  Praga
```
Every search is saved to Excel file as a new row. If the weather DIV is visible on the top of results list for your search, you extract current temperature, precipitation and wind and save them to Excel. 

| Time               | City        | Temperature | Precipitation | Wind   |
| ------------------ |:----------- |:---------- |:------------- |:----- |
| 2018/06/06 12:34:56| Katowice    |         17° |             5%| 6 km/h |

### How is it built
- Readers: to read data from files (txt, xml) - implement from Ireader
- Handlers:  to manage external application (Chrome, Google) using packages - implement from IHandler
- Exctractors: communicates via strings with ChromeHandler and ExcelHandler to transfer data
- Exceptions: custom exception DataNotDound for ChromeHandler

### To be added
Logger
