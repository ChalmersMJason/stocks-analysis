# stocks-analysis
Module 2 Challenge Submission

## Overview of Project: 
* Background - We have been asked by Steve to create an efficient method for analyzing stocks. In Module 2, VBA was utilized to identify the total daily volume and return for a given list of stock tickers in 2017 and 2018. The issue with the module 2 solution was that the code may not be scalable to more stock tickers and may not execute in as little time as possible. 

* Purpose - The purpose of this exercise is to refactor the code from Module 2 to make it run faster and collect information more efficiently. By changing the method of data collection, we hope to have the code run faster and output the same insights on stock performance. 

## Results:

[See file here for full excel file and VBA code](https://github.com/ChalmersMJason/stocks-analysis/blob/main/VBA_Challenge.xlsm)

In order to refactor this code, a ticker index, output arrays, & loops were utilized. The loops goes through all the rows in the spreadsheet for the given ticker and identifies total volume, starting price, and ending price. 

![image](https://user-images.githubusercontent.com/85259984/130324861-a304e971-cc11-4cef-8846-272731249af7.png)

We then updated the code to identify when a ticker has changed and increased the ticker index to repeat the process for the next stock. Once all that data was collected, we created final outputs that showed volume and percent change for each ticker.

![image](https://user-images.githubusercontent.com/85259984/130324951-c3add824-688a-4c54-8932-71c07feac11a.png)


As seen in the images below, refactoring this code improved the performance. The code ran between 0.6 and 0.7 faster than it did prior to refactoring.




![Outcome for 2017](https://github.com/ChalmersMJason/stocks-analysis/blob/main/VBA_Challenge_2017.png)
![Outcome for 2018](https://github.com/ChalmersMJason/stocks-analysis/blob/main/VBA_Challenge_2018.png)

## Summary: 
* Refactoring code is important as structuring code in the most efficient way will allow for both better performance for the end user and easier collaboration on the backend. The end users of the code we are writing should be able to get insights as fast as possible with instructions that are easy to follow. If we must pass our code off to another team member or someone else needs to use our file that understands VBA, the code should be written clean and efficiently so they can add to it as needed without wasting time interpreting it. 

* The original VBA script we had achieved the end result we were looking for. The advantage in that code was that it was sufficient and if this was not an ongoing need for our user it may not be worth the time to update it. The refactored code offered more readability on the backend & ran quicker, both of which are big advantages. The only disadvantage to the refactored code would be the time it took to create. We must weigh the pros and cons of refactoring prior to investing the time into doing it. 
