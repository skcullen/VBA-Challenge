# VBA-Challenge
Unit 2:VBA HOMEWORK 

Included in this repository is my solution for the Unit 2: VBA Homework for the Georgia Tech Analytics Bootcamp.

We were asked to write code in VBA that would take raw data we were given from the New York Stock Exchange and output the summaries for each individual stock. These summaries included values such as Yearly Change, Percentage Change, Total Stock Volume, Greatest Percentage Increase, Greatest Percentage Decrease, and Greatest Stock Volume.

The code I wrote for this data was tested on a smaller set of data in an excel workbook called alphabetical_testing. The final test was run on an excel workbook called Multiple_year_stock_data. Both workbooks divided the data up into several sheets within the workbook.

My code only includes one macro called stocks. This macro will calculate all the values described above as well as set up the tables that they go in, so that the user can easily understand the data, without having to go through the code and label things themselves. The macro is also designed to run through each of the sheets by itself, with a message box popping up at the end of each sheet letting the user know that that sheet is finished (this can be easily removed if the user decides that they want it to run all the way through by itself). Although the data that it was tested on is not shown in this repository, there are three images included of the final data which show the summary tables for each of the years.

The only challenge that I came across in this code was with the calculation for Precent Change. You calculate this value by dividing the Yearly Change by the value at the start of the year. The tester data revealed the obvious problem in calculating this value for certian stocks. If the intial value, used as the denominator, was equal to 0, then the value was incalculable. I had to add in an extra if statement to deal specifically with these cases where the intial value equaled 0. Since the actual output in incalculable, I changed the output so that it would reflect what the Percent Change reflected in the rest of the data set. I debated whether I should do this method or try and change the output to reflect the fact that the value is actually incalculable. I left the two lines of code that put the values (they are commented out, so they do not show unless the user needs them) for the initial value and ending value for each individual stock in order for the user to easily check these values, if they are concerned about any strangeness involving the calculation of Percent Change.

Includes:
1. the vba macro used as a .bas file
2. Screengrabs for 2014,2015, and 2016 from the Multiple_year_stock_data excel workbook
