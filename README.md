# Stock Analysis

## Summary
I analyzed the total volume traded and the percentage return on 12 different green energy stocks using Visual Basic for Applications in Microsoft Excel. There is data for both 2017 and 2018.

## Results

In general, the twelve stocks did much better in 2017 than 2018. ENPH and RUN were the only stocks that had a positive return in both years.

![All Stocks 2017](https://i.imgur.com/Ba8qLFv.png)
![All Stocks 2018](https://i.imgur.com/yjr1OQ8.png)

By using the variable "tickerIndex" I was able to make the code run much faster than in the previous document.For example, to output the data in the refactored version I used this code: 
 
       
        Worksheets("All Stocks Analysis").Activate
       Cells(i + 4, 1) = tickerS(i)
       Cells(i + 4, 2) = tickerVolume(i)
        Cells(i + 4, 3) = tickerEndingPrice(i) / tickerStartingPrice(i) - 1


It was much faster than using the original code, which was:


    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
In the original, the program has to run through the data one ticker at a time and output the data after each one. In the refactored version, because we are using an index and arrays as opposed to the name of the ticker it can run all at once and output the data at the end. Here are the differences in times:


Original Timers:

![2017 Timer1](https://i.imgur.com/1X3YLLX.png)
![2018 Timer 1](https://i.imgur.com/S8DZcAD.png)

Refactored Code Timers:

![2017 Timer2](https://i.imgur.com/a6s9GMP.png)
![2018 Timer 2](https://i.imgur.com/Ncmzhob.png)

## Summary:

### What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code are that you have a lot done for you already and it can save you a lot of time if the original code is written well. The disadvantages are that there can be pieces you don't understand because you were not the one to write it, and it can take a long time to understand what the original writer was trying to (even if it was written well.

### How do these pros and cons apply to refactoring the original VBA script?
The advantage here was that a lot of the code was already written for me and there were a lot of steps I could skip. Formatting, headers, and many variables were already done. The difficulty was figuring out the missing pieces of code and trying to understand what the original writer's intention was, spceifically with the "tickerIndex" variable. In the end, it ended up being a much more well written code than the original but it did take time to understand the meaning of the index and using it to output data to the sheet.
