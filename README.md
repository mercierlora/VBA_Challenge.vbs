# VBA_Challenge.vbs
Refactor VBA Code and Measure Performance

1. The tickerIndex is set equal to zero before looping over the rows.

Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.
![TCIKE](https://user-images.githubusercontent.com/100738861/181917553-9079b1cc-c776-426f-9f91-c228a8a753a7.png)


2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

![image](https://user-images.githubusercontent.com/100738861/181917594-33b1aa63-2166-4ef1-8c4d-825f438e1c78.png)

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

![image](https://user-images.githubusercontent.com/100738861/181917634-ca3e3521-1b09-4a2c-89d6-de2a8cc10c75.png)

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

![image](https://user-images.githubusercontent.com/100738861/181917660-0bd4e302-e146-43e0-aeea-3451d8081540.png)


Stored values from tickerStartingPrices and tickerEndingPrices

Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.

![image](https://user-images.githubusercontent.com/100738861/181917670-05db2472-7de0-41be-8054-f16f0ebd5211.png)

5. Code for formatting the cells in the spreadsheet is working.

We make positive returns green and negative returns red, to make it a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.

![image](https://user-images.githubusercontent.com/100738861/181917719-8e3eb30d-041c-43c3-b3a9-020e0f9b2f86.png)

6. There are comments to explain the purpose of the code.

Adding Comments is requiered, as a Best Practices for Writing Super Readable Code such,

Commenting & Documentation,
Consistent Indentation,
Avoid Obvious Comments.
Code Grouping,
Consistent Naming Scheme,
DRY (Don't Repeat Yourself) Principle,
Avoid Deep Nesting,
Limit Line Length, etc...
![stock ana](https://user-images.githubusercontent.com/100738861/181917761-b7695e11-2a31-4c8f-979d-b196f0ac6295.png)


7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named Dataset Examples Provided). In adition, in our resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook..

8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.

Time on VBA_Challenge_2017.PNG

![stocks 2017](https://user-images.githubusercontent.com/100738861/181917885-e37c4f45-da70-4077-83ed-4edab00247ce.png)

Time on VBA_Challenge_2018.PNG

![image](https://user-images.githubusercontent.com/100738861/181917911-b1012db7-8518-42da-a5a2-e250ffcf964b.png)

SUMMARY: Our Statement:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
Refactoring process can affect the testing outcomes.
Advantages:
Logical errors easily appear in well structure code that contains nested conditionals and loops.
Using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
