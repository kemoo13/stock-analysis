# Stock Analysis Using Excel and VBA

## Project Overview

### Purpose
In this project, we will be refactoring a code in VBA to analyze sock data. 

### Background
Our client, Steve loved the original workbook that we created for him. Now, he would like to be able to expand that dataset over the last few years. To do this, I refactored the Microsoft Excel VBA code that was written in order to collect stock data for various years. We will be using VBA in Excel in order to automate analyses so that this code may be used to analyze any stock. For the purpose of this study, we will be refactoring the code from Module 2 to loop through all of the data for the year 2017 and 2018. 


## Results

### Analysis
To begin, I downloaded challenge_start_code VBS file that provided me with the code to refactor in Excel. I then copied the code into the Visual Basic Editor. I began refactoring by creating a tickerIndex variable and setting it to zero. Three output arrays were then created; tickerVolumes, tickerStartingPrices and tickerEndingPrices. The tickerVolumes array was set as a Long data type while tickerStartingPrices and tickerEndingprices were set as Single data types. A loop was created to initialize the tickerVolumes to zero. Another for loop was then created to loop through all of the rows of the spreadsheet. Within the second for loop (j), a script was written to increase the current tickerVolumes variable and add the ticker volume for the current stock ticker. To do this, a new variable tickerIndex, was created. We then check if the current row is the first row using the tickerIndex. If this proves to be true, then we assign that variable to the current starting price. Then we check if the current row is the last row selected, then we assign this variable the current ending price. We then increase the tickerIndex is the next row’s ticker does not match the previous one. Another loop was then created to loop through the arrays in order to output the “Ticker”, “Total Daily Volume” and “Return” in the columns of the worksheet. The code has now been refactored and is ready to be run to confirm that they are the same as the example. A pop-up message will appear on the screen showing the elapsed run time for the script for each year. 



## Summary
Refactoring code proves to have both advantages and disadvantages. Most notably, refactoring code allows for automated analysis of large data sets like the one used in this project. While this is highly useful, one must be wary to make sure they do not change too much otherwise you may run into errors. Refactoring code takes a large amount of time and there were often errors that I ran into. Despite this, once the code is completed, it can be saved and used later with similar datasets. Even to improve upon your refactored code, you need to ensure that you do not alter too much. Overall, refactoring code is a great asset if one is comfortable and knowledgeable in VB script and Excel.
