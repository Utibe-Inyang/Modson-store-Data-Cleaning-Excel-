# Cleaning Modsone E-commerce store dataset

### INTRODUCTION 
Cleaning Modson store messy e-commerce orders dataset in Excel, checking it for errors, and turning it into sales insights.

### Business question: 
Which products, countries and payment methods drive revenue, and how reliable is the data behind those numbers?

Here is the workbook with sheets containing the functions and cleaned dataset [Raw and cleaned data!](E_COMMERCE%20DATASET.xlsx)

### Excel Skills used:
- Vlookup()
- IF()
- LEFT(), FIND()

### Ecommerce Dataset
The dataset used for this project contains real-world e-commerce information which includes detailed product information on:
Order ID, Order Date, Customer Name, Country, City, Product Category, Product Name, Quantity, Unit Price, Discount, Total Sales, 
Payment Method, Order Status, Delivery Date, Customer Rating

## 1. Data Cleaning
### Duplicate Values:
Using the Order ID column, Six duplicate values was identified and cleaned.

![duplicate](Screenshots/duplicate_values.png)


### Inconsistent values:
There are inconsistent data in 
- Country 
- Product category 
- Payment Method 
- Order Status


![inconsistency](Screenshots/inconsistency.png)


### Data Quality Issues
- Discount stored as 20 instead of 20%
- Discount exceeds 100
- Dates stored as text
- Category label did not match the product code
- Day and month swapped	in Order Date column
- Customer rating greater than 5
- Negative number in quantity
- Order Date should be greater than or equal to Delivery Date, but in some cases, they are not so.

   ![discount](Screenshots/discount.png)


### Missing values
- Missing rating
- Missing delivery date
- Missing customer name

  ![cus_rating](Screenshots/cus_rating.png)

  ![delivery_date](Screenshots/delivery_date.png)

  ![cus_name](Screenshots/cus_name.png)


  ## Formulas and Functions
  - I used this function to extract the first four product code from Product Name column
  ```
  =LEFT([@[Product Name]],FIND("-",[@[Product Name]])-1)
  ```

  - I use the function below to return the correct Cleaned Product Category for each extracted four letter Product code
    
    ```
    =VLOOKUP(I140,REPORT!$B$1:$C$4,2,FALSE)
         ```
   - The function below compares the Original Product Category with the Cleaned Product Category column
     ```
      =IF(F140=J140,"Match","Mismatch")
     ```
  
   - The function below checks whether the Order Date is less than or equal to the Delivery Date (Which is logically right)
     ```
     =IF(B2<=R2,"YES","NO")
        ```

### Observation

- Valid, non-cancelled orders brought in 182,974 from 142 orders, an average order value of 1,289.
- Electronics earned the most (79,635, 44% of revenue). Home & Kitchen earned the least (24,879).
- Nigeria leads with 64,325 (35% of revenue). Ghana is the smallest market (15,900).
- Sales peaked in September (21,405) and were lowest in January (9,252).
- Credit Card is the most used method (69 orders, 49%). Cash is the least used.
- 15% of all orders were cancelled and 17% are still pending, so about 32% of orders are not producing revenue yet.
- Average rating is 3.20 out of 5, and 43% of valid ratings are 4 or 5. Delivered orders took 18.2 days on average to arrive.
- 26 of 191 rows (14%) were excluded from revenue because of impossible values. The category label disagreed with the product code in 133 rows (70%)

## Conclusion
- After cleaning this dataset, I noticed some missing values and data quality issues. 
- There are 2 conflicting columns namely: Original Product Category and Product Name. I created a dummy column from Product Name by extracting the first four letters which is an abbreviation for Beauty, Electronics, Fashion and Home & Kitchen. But on a closer look, the extracted column and the Original Product Category do not match. 

So  the headache here is;  Should we go with the Product Name or the original Product category for our Analysis?

- Order Date should be less than or equal to Delivery Date, but in some cases, they are not so. 
Therefore, to get optimum insights, this data MUST be sent back to the store manager to resolve those clashes
