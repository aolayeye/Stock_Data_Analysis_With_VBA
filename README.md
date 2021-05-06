# stock-analysis
# Overview of Project
The green stocks analysis is about investigation whether an investment in DAQO New Energy corporation stocks will be a good investment. The analysis looks at DQ's stock for the year 2017 and 2018. As a way to provide further insights into how DQ compares with other stocks, and determine whether an investment in any other stock will be worth it; this stock analysis expands to all stocks for 2017 and 2018. the analysis quickly expands to include all stocks for 2017 and 2018.
The solution provided at the end of the analysis has been implemented in way that allows scaling, meaning: if we decide to analyze the more rows of stock and or more years, we would be able to do so with the click of a button. Using code to automate tasks decreases the chance of errors and reduces the time needed to run analyses, especially if they need to be done repeatedly.
This analysis will help inform the decision of which green energy stock to invest in. The goal is to use a data driven approach to decision making. we would use code to automate analysis so we can reuse the same code for any stock for any number of years. we would be able to answer the following questions:
1. Is DQ stocks worth investing in?
2. Are there other stocks as good as or better than DQ?
3. If Yes, can we diversify our investment portfolio
4. If we wanted to analyze more stocks for multiple years, will our solution be able to handle that type of scale?

## Step 1: DQ Anlaysis for 2018
The total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
we utilized loops to determine the total volume of DQ stocks traded for 2018
1. Number of rows to loop over

### Step 2: Analyze a whole list of stocks.
1. manually activate worksheet that contains stock data to analyze
2. to loop over a list of stocks, we implemented an array solution 
3. to loop through the different stocks, and the every row to determine the total volume of stocks, we implemented a nested loop 
4. we leveraged existing DQ analysis code we craeted to determine volume traded and return for the year.

### Control Flow
1. Format the output sheet on the "All Stocks Analysis" worksheet.
2. Initialize an array of all tickers.
3. Prepare for the analysis of tickers.
* Initialize variables for the starting price and ending price.
* Activate the data worksheet.
* Find the number of rows to loop over.
4. Loop through the tickers.
5. Loop through rows in the data.
* Find the total volume for the current ticker.
*Find the starting price for the current ticker.
* Find the ending price for the current ticker.
6. Output the data for the current ticker.
* static formatting
* conditionla formattiing
* color formatting
7. Run Button
8. Run analysis for any year

## Results


![All_Stocks_2017](https://user-images.githubusercontent.com/67847583/117246846-edef9080-ae02-11eb-8419-c34178f14cca.png)
![All_Stocks_2018](https://user-images.githubusercontent.com/67847583/117246860-f1831780-ae02-11eb-913c-5a328d4d34b8.png)


![Improved_Code_Logic_1](https://user-images.githubusercontent.com/67847583/117247736-73277500-ae04-11eb-9b9e-7a96cf14a0f6.png)
![Improved_Code_Logic_2](https://user-images.githubusercontent.com/67847583/117247747-77ec2900-ae04-11eb-8424-f6e7db3a3d69.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/67847583/117247811-905c4380-ae04-11eb-92b8-4e555a7e8585.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/67847583/117247816-92be9d80-ae04-11eb-81ff-f2ed4ffe2551.png)






### Code Performance
Intial performance measures shows that our code ran for 0.82 seconds for 1 year. after refactoring and added extra functionalty for our code to process all stocks at the same time and format at run time, we got our code to run in 0.78 seconds

![VBA_Challenge_2017](https://user-images.githubusercontent.com/67847583/117246156-b8967300-ae01-11eb-9ea8-470b184090c0.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/67847583/117246172-c0561780-ae01-11eb-88c3-3e4545cbbcf6.png)

## Sumary
