# Refactored Stock Analysis 

## Overview of Project

The purpose of this project is to take Steve's current workbook on stock data and see if we can make
it run more efficiently by refactoring the code. We will compare our refactored method versus the prior
method and identify the pros and cons of refactoring code and see how that relates to adjusting scale
of a data set. 

## Results

In our original method of analyzing this stock data, we had the code line up in the form of a nested loop.
After is prompted then enters a `yearValue`, the original code would take each stock present in the file
scan every row of stock data, and gather values for stock volume, and start/end prices. Once the script
sclaes the entire file to pull info about that stock value, it then repeats the process for the other stocks
listed in our array until all array elements are covered. In performing this method, the same large set of 
data is scanned multiple times, equal to the number of elements in our array. This can be demonstrated by the
abridged example code below.  

Original Method
```
For i = 0 To x
	
	ticker = tickers(x)
	`Initialize values
		
	For j = 2 To RowCount
		`Read all the rows and gather info to relevant ith index
			
	Next j
		
Next i```

#### Runtime for Original Method on 2017 data
!Resources/VBA_Challenge_Original_2017.png

Runtime for original method on 2017 data: 1.4961 seconds

#### Runtime for Original Method on 2018 data
!Resources/VBA_Challenge_Original_2018.png

Runtime for original method on 2018 data: 1.5195 seconds

Take note of these runtimes from the original method. We will compare them to our Refactored results momentarily.

When we refactored our code we came in with the goal of having the code run more efficiently. To do this, 
we decided to have the code run through the workbook only one time, reading all information about the stocks
present without having to do repeated iterations through the workbook per stock ticker present in our array. 
What made this possible is that our data is already grouped up as such that the stocks are all grouped together
in consecutive rows which made it possible to better write code and gather data about the new stock data by 
simply reading consecutive rows. When a group of stocks were read, we had an index value, `tickerIndex`, that 
would increment and start gathering data about the next stock in our file. This is demonstrated by the abridged
refactored example below.

Refactored Method
```
`Initialize tickerIndex variable
tickerIndex = 0

For i = 2 to RowCount
	`Read rows to determine starting/current rows of current stock data gather stock data 
	
	`Read final row of current stock data...
	If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i,1).Value = tickers(tickerIndex) Then
	
		tickerEndingPrices(tickerIndex) = Cells(i,6).Value
		
		`Then onto the next stock!
		tickerIndex = tickerIndex + 1 
	
	End If
	
Next i```

To have this work effectively, we have to assume that the data is lined up as such that all elements of a stock 
within the workbook are grouped together and in ascending date order.

#### Runtime for Refactored Method on 2017 data
!Resources/VBA_Challenge_2017.png

Runtime for refactored method on 2017 data: 0.15625 seconds

#### Runtime for Refactored Method on 2018 data
!Resources/VBA_Challenge_2018.png

Runtime for refactored method on 2018 data: 0.16406 seconds

As we see, in comparing the Original method versus the Refactored method, the refactored code runs over 10 times faster
than our original method. While we gain efficiency in processing time, again our assumptions listed earlier have to be in 
place for this to be a viable method. 

## Summary

- What are the advantages and disadvantages of refactoring code?
In refactoring code, we gain a strong performance advantage. As observed in our Results section, the refactored method
grabbed us our results over 10 times more quickly than our original method! In performing analysis about many different 
stocks, the performance efficiency gained could save our user time in the short term and get to investing more quickly!
However, in our performance gain, we have to have our information lined up in a specific way to have things work properly. 
What we gain in performance, we lose in flexibility on how we should line up our worksheet. 

- How do these pros and cons apply to refactoring the orignal VBA script? 
When applying the earlier summary to our immediate example with the VBA script, the performance gained was instantly seen 
when we refactored our code. However, for our code to work as expected, we had to make a TON of assumptions about our 
worksheet and how our data lined up. All the stock information was grouped and sorted properly all while within each stock 
block, the entries were sorted chronologically in ascending order. Without these sorting elements in place, we would run into
overflow issues with our code. For example, if our stock entries weren't grouped by stock and they were scattered throughout 
the sheet, our `tickerIndex` used in the refactored method would exceed the number of elements in our array and would blow up
our code when we attempt to run it. Further, without proper chronological sorting, we would have a high probability of having 
inaccurate return values because our starting and ending price values would be different. While the refactored code ran more 
quickly, the original method compensates by continuing to properly line up our information regardless of how our stock entries
are sorted. Even then, we would still have to consider further logic in our original code method to ensure we had the absolute 
earliest value recorded for our starting price, and our latest value recorded for ending price within the stock as it runs 
through all of the rows. 