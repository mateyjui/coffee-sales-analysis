# Coffee Shop Sales Analysis - Excel
Excel-based coffee shop sales analysis using XLOOKUP, INDEX-MATCH, IF, Nested IF's, Pivot Tables, and an interactive dashboard.

## Objective
As part of strengthening my Excel skills, I worked on a coffee shop sales analysis project using Orders, Products, and Customers tables.

## Dataset
1,000 transactions across Orders, Products and Customers tables.

## Excel Skills Used
- XLOOKUP
- VLOOKUP
- INDEX-MATCH
- IF
- Nested IF
- Pivot Tables
- Slicers
- Timeline
- Charts

## Data Preparation
### Data Cleaning
* Cleaned the dataset by identifying and removing duplicate records using the **Remove Duplicates** feature.

### Populating Customer Details
* Applied **XLOOKUP()** using `Customer ID` as the lookup value to populate **Customer Name, Email, Country, and Loyalty Card** details from the `Customers` sheet.
  ```excel
  =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)

* Used **IF()** in combination with **XLOOKUP()** to return a blank instead of `0` for customers with missing email addresses.
  ```excel
  =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0))

### Populating Product Details
* Applied **INDEX-MATCH** to retrieve **Coffee Type, Roast Type, Size, and Unit Price** from the `Products` sheet.
* Used **MATCH()** to identify the position of the selected `Product ID`, while **INDEX()** retrieved the corresponding product attribute.
* Used **absolute references (`$`)** to keep the lookup ranges fixed when copying the formulas across rows and columns.
  ```excel
  =INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
  
### Additional Calculations and Transformations
* **Sales:** Calculated sales by multiplying **Unit Price × Quantity**.
* **Coffee and Roast:** Used **Nested IF()** statements to convert abbreviated coffee types into their full names.
  ```excel
  =IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))

## Dashboard
The Dashboard highlights sales trends for the period 2019-2022, country-wise sales for Ireland, Uunited Kingdom and United States, and identifies the top 5 customers.

## Files
Orders: Contains the coffee shop transactions. Each row represents an order line for a particular product purchased by a customer.
Products: Contains information about the coffee types available for sale. It is used to identify product characteristics and pricing.
Customer: Contains customer-level information used to identify the customers associated with each transaction.

## Key Learnings
This project helped me understand how Excel functions can be combined for data preparation, and how Pivot Tables and dashboards can turn raw transactional data into meaningful business insights.
