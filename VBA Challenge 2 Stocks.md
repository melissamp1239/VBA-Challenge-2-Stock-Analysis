# VBA Module 2 Challenge

## Overview of Project
Steve wants to analyze the entire stock market over the past few years for his parents.  Previously, he analyzed a data set of just twelve lines and now he'll need to analyze a data set of several thousand lines, thus he needs to refactor, or simplify his code to have it run more quickly.

### Purpose/Background
Instead of reviewing only a dozen stocks, Steve wants to expand his analysis to the entire stock market over the last few years.  However, using the same code to go through twelve lines of data may take longer if you go through thousands of lines of data. Thus the code needs to be simplified or refactored. When refactoring code, you are not adding new functionality.  You are simply making the code more efficient by taking fewer steps.
## Results

The refactored program starts with the question, What year would you like to run the analysis on?

**The following is an an example of the code used to ask this question:**
` yearValue = InputBox("What year would you like to run the analysis on?")`

The refactored code tended to run faster in 2017 (0.06252018 seconds) than in 2018 (0.0859375). 

**2017 Code Timing**
![This is an image](https://github.com/melissamp1239/VBA-Challenge-2-Stock-Analysis/blob/main/2017refactored_time.png)

**2018 Code Timing**
![This is an image](https://github.com/melissamp1239/VBA-Challenge-2-Stock-Analysis/blob/main/2018_refactored_time.png)

Overall,the annual returns for the tickers measured were better in 2017 than in 2018. The stocks with the best annualized returns in 2018 were ENPH (81.9%) and RUN (84.0)%. Interesting, the annualized rate of return for ticker ENPH was actually higher in 2017 (129.5%) compared to 2018 (81.9%).

**Challenges

I found the following code requiring moving through the tickers challenging to write, but it helped the program move quickly through the data:

**Code to check if the current row is the first row with the selected ticker index**

`If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value` `= tickers(tickerindex) Then`
`'AssignCurent stock price to ticker Starting Prices Value`
 `tickerStartingPrices(tickerindex) = Cells(i, 6).Value`

 `End if`

 **Code to check if the current row is the last row with the selected ticker and if it is , then assigning the current closing price of the tickerEndingPrices variable**

  `If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then`

` tickerEndingPrices(tickerindex) = Cells(i, 6).Value`
            
`tickerindex = tickerindex + 1`

`End If`
    
   ` Next i`

## Summary

### Advantages and Disadvantages of refactoring code
The source of advantages and disadvantages comes from the following: [(https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)]

According to Martin Fowler (Father of Code Smell), the **advantages** and disadvantages of refactoring are:
    1. Refactoring improves the design of software
    2. Refactoring makes software easier to understand
    3. Refactoring Helps Findig Bugs
    4. Refactoring Helpfs Programming Faster

**Disdvantages** include:
    1. Run out of money
    2. Run out of time

### Pros and Cons of the original and refactored VBA Script

The original DQ Analysis code would require much more voluminous code to put in the individual tickers to loop over all of the rows and tickers.  Instead, we refactored that part of the code with the code I list under the *Challenge* section of this document. This is much more concise code.

Additionally, you have to know how long your list is with the original code. 

`rowStart = 2`
`'DELETE: rowEnd = 3013`
`'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists`
`rowEnd = Cells(Rows.Count, "A").End(xlUp).Row`
`totalVolume = 0`

Additionally, via a popup box, the refactored code program asks the user what year they want to look at.  However, for the original code, you would have to specify a new year in the code.

If you did not have much time to refactor your code, I think that patterning your code after the the original code would be good.  The original code is much easier for a new programmer to understand.








        
