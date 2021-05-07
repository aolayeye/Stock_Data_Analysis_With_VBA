# stock-analysis
# Overview of Project
The green stocks analysis project is about investigating whether the DAQO New Energy corporation stock will be a good investment. The analysis looks at DQ's stock for the years 2017 and 2018. TThis stock analysis expands to all stocks for 2017 and 2018 to provide further insights into how DQ compares with other energy stocks and determine whether an investment in any other stock will be worth it. The solution provided at the end of the analysis has been implemented in a way that allows scaling. If we decide to analyze the more rows of stocks for different years, we will do so with the click of a button. Using code to automate tasks decreases the chance of errors and reduces the time needed to run analyses, especially if they need to be done repeatedly. The goal is to use a data-driven approach to decision-making. We would use code to automate stock analysis, and we will reuse the same code for any stock for any number of years. We would be able to answer the following questions:

1. Is DQ stocks worth investing in?
2. Are there other stocks as good as or better than DQ?
   - If Yes, can we diversify our investment portfolio
4. If we wanted to analyze more stocks for multiple years, will our solution handle that type of scale?

### Step 1: DQ Analysis for 2018
Our first step is to determine the total daily volume and yearly return for the DQ stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference between the starting price and the ending price.
We utilized loops to determine the total volume of DQ stocks traded for 2018

### Step 2: Analyze a whole list of stocks for 2017 and 2018
We expanded our code to analyze multiple stocks with the added flexibility of selecting which year we want to analyze.
1. We dynamically activated the worksheet that contains stock data to analyze
2. To loop over a list of stocks, we implemented an array solution 
   - To loop through the different tickers and every row to determine the total volume of stocks, we implemented a nested loop 
4. We leveraged the existing DQ analysis code we created to determine volume traded and return for the year.

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
   * Find the starting price for the current ticker.
   * Find the ending price for the current ticker.
6. Output the data for the current ticker.
   * static formatting
   * conditional formatting
   * color formatting
7. Run Button
8. Run analysis for any year

## Results
The general trend for the energy stocks for the two years analyzed shows that 2017 was a generally positive year in contrast to 2018, where all stocks except ENPH and RUN had a negative return on investment.
While DQ made a 199.4% return in 2017, it made a 62.6% loss in 2018. We may need to analyze more years for DQ to determine the trend.
An initial recommendation would be ENPH and RUN since they show a consistent trend for the years analyzed.
In general, two years' worth of stock data is too small to decide which stocks to pick for investment.

![All_Stocks_2017](https://user-images.githubusercontent.com/67847583/117246846-edef9080-ae02-11eb-8419-c34178f14cca.png)
![All_Stocks_2018](https://user-images.githubusercontent.com/67847583/117246860-f1831780-ae02-11eb-913c-5a328d4d34b8.png)

While we began our analysis for the DQ stock, we refactored our code to achieve:
  * scale: if we decided to analyze more years for more stocks in the future
  * speed: our refactored code ran in 0.20 seconds less time than the original code
  * maintenance: since we have a small codebase, refactoring was an easy exercise, making future maintenance activities easier to implement.
  
### Improved Code Logic
![Improved_Code_Logic_1](https://user-images.githubusercontent.com/67847583/117247736-73277500-ae04-11eb-9b9e-7a96cf14a0f6.png)
![Improved_Code_Logic_2](https://user-images.githubusercontent.com/67847583/117247747-77ec2900-ae04-11eb-8424-f6e7db3a3d69.png)

### Improved Performance
#### Original Code
![Original_Code_Performance_2017](https://user-images.githubusercontent.com/67847583/117377255-766c4080-ae98-11eb-8f42-2b22fb79d6e8.png)
![Original_Code_Performance_2018](https://user-images.githubusercontent.com/67847583/117377259-79ffc780-ae98-11eb-9729-e334a6a8d442.png)

#### Refactored Code
![VBA_Challenge_2017](https://user-images.githubusercontent.com/67847583/117247811-905c4380-ae04-11eb-92b8-4e555a7e8585.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/67847583/117247816-92be9d80-ae04-11eb-81ff-f2ed4ffe2551.png)

Initial performance measures show that our code ran for 0.97 seconds for one year. After refactoring and adding extra functionality for our code to process all stocks simultaneously and format results at run time, we got our code to run in 0.79 seconds.

## Summary

### Refactoring

Refactoring is the process of restructuring existing code without changing its external behavior. Refactoring improves the design, structure, and implementation of the software (its non-functional attributes) while preserving its functionality. 

#### Advantages of Refactoring
When code is refactored as standard practice from when you begin writing, multiple advantages are realized:
1. Maintenance: Since refactored code is dynamically written and well commented and documented, it reduces complexity, and the code is easier to read and understand. Thus down the line, code is easier to maintain, and new functionality can be easily added. 
2. Scale: It is easier to scale refactored code; case in point, if we refactored our code to use dynamic arrays or used a list method to determine that number of tickers, then we would never need to worry about:
   1. analyzing more than 12 tickers
   2. we would not need to manually initialize our ticker list
3. Speed: Generally, refactored code achieves improved speeds when the code base is not too large
4. Reuse: When code is refactored correctly, code can be reused
5. Removes Code Smells: Refactoring removes duplicate code, long classes or methods, variable mutations, contrived complexities, data clumps, dead code, etc. By removing code smell, we are potentially preventing future bugs.
#### Disadvantages of Refactoring
1. Since refactoring often requires extracting code structure, data models, and intra-application dependencies to get back knowledge of an existing software system; it can be a time consuming and expensive process
2. Chances for costly mistakes are high since refactoring code usually requires in-depth knowledge of existing code

#### Refactoring the VBA Challenge
1. One clear advantage of refactoring the stock analysis code is our ability to process multiple stocks simultaneously. We also refactored our code to give us the flexibility of specifying the year we wish to analyze.
2. Our code is well commented; this makes it possible to follow through, understand and maintain the code in the future.
3. We adopted a dynamic approach to determine the number of rows to loop through, and we utilized a nested loop structure that would allow our code to scale easily in the future.
4. By refactoring our code, we shaved 0.20 seconds from the initial code.

If we want to get the full benefits of refactoring, my submission is that:
1. Refactoring should be a standard coding practice, incorporated from day one
2. We should give careful thought to the question of "when to refactor." The following considerations also require careful thought:
   - If the chances for enhancements are high, it may be a good idea to refactor your code beforehand
   - Sometimes, bad patterns like tight coupling, duplicate code, long methods, large classes, etc., are detected in the code. This scenario will be an excellent case to justify refactoring.
   - Fixing too many bugs could result from code smells; in this case, refactoring will be a preferred option rather than bug fixing.

3. It is also important to note that we do not want to refactor when the code is stable.
4. Refactoring should not be done when the cost of refactoring is much higher than the cost of rewriting from scratch  
