# Stock Analysis Using Excel and VBA

## Project Overview

### Purpose
In this project, we will be refactoring a code in VBA to analyze sock data. 

### Background
Our client, Steve loved the original workbook that we created for him. Now, he would like to be able to expand that dataset over the last few years. To do this, I refactored the Microsoft Excel VBA code that was written in order to collect stock data for various years. We will be using VBA in Excel in order to automate analyses so that this code may be used to analyze any stock. For the purpose of this study, we will be refactoring the code from Module 2 to loop through all of the data for the year 2017 and 2018. 


## Results

### Analysis
To begin, I downloaded challenge_start_code VBS file that provided me with the code to refactor in Excel. I then copied the code into the Visual Basic Editor. I began refactoring by creating a tickerIndex variable and setting it to zero. Three output arrays were then created; tickerVolumes, tickerStartingPrices and tickerEndingPrices. The tickerVolumes array was set as a Long data type while tickerStartingPrices and tickerEndingprices were set as Single data types. A loop was created to initialize the tickerVolumes to zero. Another for loop was then created to loop throug all of the rows of the spreadsheet. Within the second for loop (j), a script was written to increase the current tickerVolumes variable and add the ticker volume for the current stock tiker. To do this, a new variable tickerIndex, was created. We then check if the current row is the first row using the tickerIndex.



## Summary

