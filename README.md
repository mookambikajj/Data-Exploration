# Product Data Exploration Using Excel

A simple data analytics project using Microsoft Excel to explore and analyse product data.

## Description

This project analyses a product dataset with 34 records and 10 columns.
Excel functions like SUM, AVERAGE, MIN, MAX, IF, SUMIF, COUNTIF, LEFT, RIGHT, and MID
are used to find useful insights from the data.

## Getting Started
Dataset
Column NameDescriptionProduct IDUnique code for each productDayDay taken from Product IDCountry CodeShort code for the countryMonthMonth taken from Product IDProduct NameName of the productBrand NameBrand that makes the productPrice ($)Price of the product in dollarsQuantityNumber of units availableCategoryType or group of the productPrice RangeWhether the product is Low, Medium, or High price

### Dependencies

* Windows 10 or above
Microsoft Excel 2016 or later
Basic knowledge of Excel.
  
### Installing
*Download and open Excel Assignment 1__Data_Exploration.xlsx in Excel
*Go to the Dataset sheet to view the raw data
*Go to the Analysis sheet to see the formulas used
*Go to the Dashboard sheet to view charts and insights

### Executing program
Transformation
Function
Output
Extracted Day from Product ID 
=LEFT(A2,2)
28
Extracted Country Code from Product ID 
=RIGHT(A3,2)
US
Extracted Month from Product ID 
=MID(A2,4,3)
JAN
Classified products by price level 
=IF(G2>=500,"High Price", "Standard Price")
High Price / Standard Price

1. Descriptive Analysis (What happened?) 
Summarises the overall distribution and key statistics of the product data.

Metric 
Function
Result
Total Price of All Products 
=SUM(G2:G35)
                         $ 10,100
Total Number of Products 
=COUNTA(E2:E35)
                                    34
Average Product Price 
=AVERAGE(G2:G35)
                               $ 297
Minimum Price 
=MIN(G2:G35)
                                $ 30
Maximum Price 
=MAX(G2:G35)
                            $ 1,000
Total Price — Electronics 
=SUMIF(I2:I35,"Electronics",G2:G35)
                            $ 8,050
Products Priced Below $100 
=COUNTIF(G2:G35,"<100")
11

## Help

#VALUE! → Make sure Price and Quantity columns have numbers, not text
#REF!   → Check that no rows used in formulas have been deleted
#NAME?  → Check spelling of function names (e.g. COUNTIF not COUNIF)

## Authors
Mookambika
Course: Data Analytics

## Version History
    * v0.2 — Final version with all analysis and dashboard completed
    * v0.1 — Initial version with basic dataset and formulas
    
## License

Microsoft Excel Documentation — https://support.microsoft.com/excel
Data Analytics Course Materials

