# Stock Analysis using VBA

## Overview of Project

In this project, we are helping Steve analize some stock data for his parents. Steve wants to find the total daily volume and yearly return for each stock.
A workbook that analyzes stocks at the click of a button has been created for Steve. However, he now wants to expand the dataset to include the entire stock market and our current VBA code might not be capable or efficient enough to do so. 

![Link to Stock Analysis file](VBA_Challenge.xlsx) 

### Purpose
The purpose of this challenge is to edit or refactor the original code we have prepared for Steve to make it more efficient. We want the code to provide us with the same information as the original code by looping through all the data once. After refactoring, the code should take fewer steps, using less memory, or the logic of the code is improved to make it easier for future users to read. 

## Results
To refactor the original code, the *tickerIndex* is set to zero before looping over the rows and the arrays were created for *tickers*, *ticketVolumes*, *tickerStartingPrices*, and *tickerEndingPrices*

![code1](/images/code1.png)

To access the stock ticker index for the *tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices* arrays,  the *tickerIndex* is used.
The new script loops through stock data, reading and storing all of the following values from each row: *tickers, tickerVolumes, tickerStartingPrices*, and *tickerEndingPrices*.

![code2](/images/code2.png)

After running the refactored code, both the outputs, original and refactored, for the 2017 and 2018 stock analyses match. <br />
Here is the output for 2018 original and refactored code: 
<br />
![outputrefactored](/images/outputrefactored.png)
![outputroriginal](/images/outputoriginal.png)

The execeution times of the original script and the refactored script are:<br />
For 2017:
<br />
![orignal2017](/Resources/VBA_Challenge_2017.png)
![refactored2017](/Resources/VBA_Challenge_2017_Refactored.png)
<br />
For 2018: 
<br />
![original2018](/Resources/VBA_Challenge_2018.png)
![refactored2018](/Resources/VBA_Challenge_2018_Refactored.png)

The refactored code is much more efficient and it runs approximately 0.9 seconds faster than the original script.


## Summary
### Advantages or disadvantages of refactoring code

The advantages of refactoring code is that it improves the design of the code. It makes the code  simpler, cleaner and more organized. 
By doing so, it is much easier for other users to understand and make alterations when necessary. 
Refactoring can also be a great way to find bugs and reduce errors in the code since we are re-organizing and revisiting the code. People usually start spotting bugs that might have been missed before and improve the quality by fixing them.
Lastly, after refactoring, code usually runs faster since it problably takes fewer steps and uses less memory than the original code. 

 
The disadvantages of refactoring is that it takes time. In real life projects, it can slow down or delay the development process. 
Refactoring can be expensive since more time has to be spent on it and it increases cost of projects.
If not done properly, modifying code can introduce new bugs and errors into the code and the final output might not be the same as the original. 

### Pros and cons of refactoring the original VBA script

The refactored code for this project takes less steps and it is easier to understand. The refactored scripts run a lot faster than the original. 
Steve can now analyze more stocks with the improved spreadsheet. 
Refactoring does take time, when working on this VBA_Challenge, it was a little bit challenging to get the code run correctly. 


