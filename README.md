# Analyzing Stock Date with VBA

## Overview of Project
Excel workbook holding stock data including, daily volume and close rates were organized by ticker and analyzed for years 2017 and 2018. Using VBA, specific stock data for ticker "DAQQ" were analyzed to calculate the yearly volume and percentage difference for the years 2017 and 2018 in one worksheet. All stock tickers were then analyzed to calculate the yearly volume and percentage difference for the years 2017 and 2018 in another worksheet. This was orginally analyzed in VBA using an array to collect the ticker data, then creating output data in a separate worksheet. Using VBA, green_stock_analysis.xlsm, stock data was analyzed and buttons were created to allow for user year input and analyze and format stock data. Using timer functions, time elapsed was calculated for run time of macro analyzing all stock analysis for the years 2017 and 2018. 

The code was then refactored in VBA_Challenge.xlsm file to monitor efficiency, using multiple arrays to loop through tickers, calculate yearly volume, and calculate percentage difference. Timers were set to collect time at beginning and end of macro to produce message box with calculated run time for each year analysis for the refactored code. Run times for 2017 and 2018 were compared for All Stock Analysis refactored and also compared to original code. 

### Purpose
The purpose of this project was to use VBA to analyze stock data in Excel and refactoring code to compare effeciency using run time analysis. 

## Results
The original code to analyze stock data for 2017 and 2018 ran in 0.8398438 and 0.7753906 seconds, respectively. The refactored code to analyze stock data for 2017 and 2018 ran in 0.15625 and 0.2109375 seconds respectively. The refactored code ran faster for both analysis years, compared to original code. 
Below is the image of the message box displaying the run time for years 2017 after refactoring code. 
![2017_Refactored_RunTime](/Resources/VBA_Challenge_2017.png)

Below is the image of the message box displaying the run time for years 2018 after refactoring code. 
![2018_Refactored_RunTime](/Resources/VBA_Challenge_2018.png)

## Summary
 
 - What are the advantages or disadvantages of refactoring code?
  - Advantages of refactoring code are that is much faster to run and usually takes up less space. Refactoring code can also make code easier to read for developers and to be used in future projects. 
  - Disadvantages of refactoring code, is that one could argue the time it takes to refactor the code, will never be made up in the milliseconds of time it saves during run times. 

- How do these pros and cons apply to refactoring the original VBA script?
 - The increased efficiency of the code was made evident by the decreased elapsed time in running the macros in VBA. The code of the refactored macro also appears simpler with the use of arrays to loop through data as well as looping through output data associated with each stock ticker. 
 - As a new programmer, the increased time it takes to re-write code is a major disadvantage. However, it helps to go through the process to become faster and write code more efficiently for future projects. 
