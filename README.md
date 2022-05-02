# stock-analysis
A VBA project to do stock analysis

### Overview of Project 

In Module 2 we learned how to perform stock analysis on different green energy stocks using subroutine and macros in VBA excel to help 
Steve make decisions for his parents about investments.Analysis was done for different tickers for a particular year(2017 or 2018)
provided at the input.The purpose of this analysis was to find out “Total Daily Volume” and “Yearly Return” for each 
stock in a given year and a get a perspective of stock performance at a glance with help of proper formatting.
The year was dynamically chosen by the user through an input box. Comparison between time execution of refactored code 
and original code was made.

### Analysis

1. Time taken for Stock analysis for the year 2017 for original & refactored code

![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Original%20code%202017.png)
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)


2. Time taken for Stock analysis for the year 2018 for original & refactored code
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/original%20code%202018.png)
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

In original code ,we use nested for-loops. First in the beginning loop tickers(i) to get the value of the current ticker. In the next for-loop we loop through each row (from 2nd row till the end of the data) to get the rest of the data and then we go to beginning of the first loop for the next ticker. So each ticker goes through all rows i.e around 3013 rows and then again same for next ticker in all 12*3013(about 36,000) rows taking it long for executing the data to determine the totalVolume and Returns.

Below is the snippet of the original code -

![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Original_code.png)

In Refactored code the analysis was done for year 2017 and 2018 for all stocks using only one for-loop to is used as we are accessing the ticker value using the tickerIndex and then for each tickerIndex we get the values for all the columns which are defined as arrays belonging to each ticker - tickerVolume, tickerStartingPrices & tickerEndingPrices. If the current ticker value is not same as the next row ticker , the ticker index is incremented and in the same loop we get the data, process is repeated for all the tickers in the array and only once in the for loop so in all only 3013 rows are executed only once. Thus, the execution time taken is shorter than in original code.

Below is the snippet of the original code -
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Refactored_code.png)


### Results

![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Stock_analysis_2017.png)
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Stock_analysis_2018.png)


The Stock results for 2017 & 2018 are as shown above. The output is same for both the original_code & refactored code.
In 2017 all stocks performed really well , with a positive return at the end of year except for TERP with only 7.2% return , which did not do well.  DQ-199.4%  was top on the list with its  returns , followed by SEDG & ENPH with 184.5% & 129.5% respectively. 

Difference between the Original_code and Refactored_code  - 
There is a huge difference between the execution times of both codes for each year. 
![](https://github.com/shrutiramana/stock-analysis/blob/main/Resources/Tabular_result.png)


### Summary
    * There is a detailed statement on the advantages and disadvantages of refactoring code in general.

The intention behind refactoring is simply to achieve better code.Following are the advantages and disadvantages of the refactoring the code in general 
	### Advantages
        * Code can be restructured without altering its functionality. 
        * With refactoring redundancies and duplications can be removed and that can improve the performance of code 
	
	### Disadvantages
        * If refractor is not done properly we can introduce new bugs and errors in the bug.
        * Since the output of the code remains the same, unless we have some evident performance improvement statistics it’s difficult to understand why 		refactoring of code is required.
	Source - https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/
	
 * There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script.

	### Advantages 
        * The code looks much cleaner with the introduction of arrays and tickerIndex and comprehendible. 
        * So instead of 2 for-loops , we use tickerIndex and only 1 for loop and we do not repeat the same process for each ticker and hence reduce redundancies & duplication.
	### Disadvantages	
        * It’s important to understand why and how we need to refactor the code, as if that is not clear it can cause confusion and might not give desired output and also can introduce bugs.
        * For a customer since the output of the code is same unless we give tangible results (like performance improvement as in the Challenge above) it would be hard for them to get convinced them for refactoring of code.

    * Teammates - Ashely Rock, Nick Foley & I collaborated for this project.

