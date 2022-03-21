# STOCK_ANALYSIS

## OVERVIEW
	Steve a recent graduate in finance has his first clients interested in investing into green energy stocks. So he has decided to take a deep look into a few green energy stocks data in order to find the stocks that would be profitable.

### Purpose
	In this analysis we will be creating an excel file, use code to find information needed and finally conduct an analysis on a few number of stocks. 

## RESULTS

### Stock performance 2017
	The analysis found that in the year 2017 a crushing majority of the stocks that were analysed made a good return on investment with profit ranging from 1 to 3 digits and only one stock plunged in a one digit negative
	
	Below is a screenshot of the 2017 results
	![VBA_Challenge_2017] ![image](https://user-images.githubusercontent.com/97865472/159197020-1cc52a65-2140-4b2e-ad78-4e55ff7a75bf.png)

### Stock performance 2018
	The analysis also found that in the year 2018 all but two stocks fell below the negative return on investment down to -62 % and only two stocks remained profitable up to 84% return on investment
	
	Below is the illustration
	![VBA_Challenge_2018] ![image](https://user-images.githubusercontent.com/97865472/159197070-cd6208ed-b452-4ff2-816a-9156d2ad8f52.png)

### Comparition of stock performance
	Compared to 2017, the year 2018 was far less profitable for the vast majority of the stocks that we analysed with only two stocks remaining above 80% of return on investment.
	It is also worth noting that those two (ENPH and RUN) had a significant change in either direction with the former going from 129.5% of profit in 2017 down to 81.9% in 2018 and the latter going from 5.5% in 2017 up to 84% in 2018.
	
	Using the following code:

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

	And also:

	  For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

	We were able to find the total volume as well as the return for each ticker.	

	Here is a screenshot of both years performance
	![VBA_Challenge_2017] ![image](https://user-images.githubusercontent.com/97865472/159197040-acde3c6e-efda-4621-b2e8-62e953da4ac9.png)
	![VBA_Challenge_2018] ![image](https://user-images.githubusercontent.com/97865472/159197087-fc8d712c-e4f2-43ba-84f8-44fdba78fffb.png)

### Comparition of execution time
	When running, the original script seems to take more time to execute compared to the refactored script.

	![2017_stock_performance] ![image](https://user-images.githubusercontent.com/97865472/159196910-6550d79d-2e0c-49bd-a8f2-0a4b7fe53ab9.png)
	![VBA_Challenge_2017] ![image](https://user-images.githubusercontent.com/97865472/159197045-9a452bc1-7ee0-42d0-a567-59387095630a.png)
	![2018_stock_performance]![image](https://user-images.githubusercontent.com/97865472/159196988-99ba1e9d-90a5-4f77-9f25-388249cca45d.png)
	![VBA_Challenge_2018] ![image](https://user-images.githubusercontent.com/97865472/159197091-59e34165-3dc1-45a4-96bb-8788fee5c20a.png)
	
## SUMMARY
	1. The advantages of refactoring code are:
		- It makes the code run faster
		- It makes the code shorter and easier to understand with fewer steps

	   The main disadvantage is it will take you time to rethink how to refactor 

	2. The advantages of refactoring our original VBA:
		- The refactored subroutine took few different ones from the original script and made them into one
		- That made the code run faster

	   The disadvantages we encoutered:
		- Too long of a subroutine creating potential identation issues
		- Risks of making mistakes or getting confused within certain lines of the script
