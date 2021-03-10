# Wall Street VBA Analysis

## *Purpose*

The purpose of this project is to analyze stock data for green energy companies in 2017 and 2018.  We calculate both the sum of the total daily volumes for each year and the annual return for each ticker symbol.

## *Results*

Overall, these green energy stocks performed much better in 2017 than 2018 with the exception of TerraForm Power (ticker symbol TERP) as you can see from the data below.

![](resources/VBA_Challenge_2017.png)
![](resources/VBA_Challenge_2018.png)

The refactored scrips ran the subroutines in 0.55 s and 0.57 s for the years 2017 and 2018 respectively.  This is a marginal improvement on 0.64 s and 0.67 s for the original script, screenshots are provided below.

![](resources/VBA_orig_2017.png)
![](resources/VBA_orig_2018.png)

### *How do the macros differ?*

The major difference in the two methods used to produce the same data is the data storage method.  In the original subroutine, cells in the worksheet are updated with each iteration of a for loop, similar to the below code.
```
For i in tickers

...

    Worksheets("AllStocksAnalysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

...

End i
```

The script for the refactored method stores the loop output data in three data arrays, which are then used to update the worksheet with the data after the data acquisition loops have run, similar to the below code.

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

...

'For loops populating arrays

...

For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

End i

```

This manipulation may seem trivial, but it did cut roughly a tenth of a second off of the run time for each subroutine.

# *Summary*

## *Advantages and Disadvantages of Refactoring Code in General*

Refactoring code may have advantages or disadvantages depending on the situation.  In any case, if the original code is not sufficiently commented, or the purpose of the code is something one isn't intimately familiar with, it may be very difficult to refactor because one may not have sufficient knowledge of the coding methods or the purpose of the code.

If code is not sufficiently commented, it may be extremely difficult for one to figure out the purpose of each statement and function.  This would make the code very difficult to refactor.

If the purpose of the code is something one isn't familiar with, it is difficult to refactor the code without knowing what the end goal is.  Take as an example electrodynamic simulations of metal semi-conductor nanoparticles using Generalized Multiparticle Mie Theory.  See the issue? No matter how good one may be at retooling code, this subject requires advanced knowledge of both electrodynamics and advanced partial differential equations. 

Furthermore, refactoring code could end up being a lot of work for a little payoff.  Take our stock analysis with VBA as an example.  If the data sets one is working with remain on the order of magnitude we manipulated here, retooling the code so that it may run 0.1 s faster may not be something one is interested in.  However, if we were to analyze larger sectors of the stock market or hedge fund performance, this retooling could be extremely valuable when the code starts taking thirty minutes or more to execute.

Refactoring code does have many advantages as well.  In general, refactoring code often makes the code more elegant, efficient, and sometimes even easier to understand.

If one is able to manipulate the code such that the same goal is achieved with fewer lines and more powerful functions, the code will generally be easier to understand and will likely execute faster as well.

Also, refactoring code gives one a deeper understanding of what the code was originally meant to do and how the end goal was achieved.  When refactoring code, one gets to take a second look at a script and decide what each statement does and if or how the code can be made more efficient.  This increase in knowledge of the code will likely mean one's commenting on the refactored code will be more useful and descriptive the second time around.

## *Advantages and Disadvantages of Our Original and Refactored VBA Script*

The most obvious advantage to our refactored VBA script is the run time.  As stated before, the refactored script runs about a tenth of a second faster than the original script.

Our refactored script also contains much better commenting so that we may look back on the code later and understand the refactored script much better than we might understand the original script.

Furthermore, our refactored script contains a one-subroutine-does-it-all macro, whereas the original script accomplished the same tasks with multiple subroutines.

Lastly, the refactored script is one that can be more readily modified to perform similar tasks for other analyses.  The data is stored in arrays in the macros themselves rather than populating cells in the worksheets while the loops run.  If we were to retool the original script for another purpose, we would almost have to completely rewrite the loops.  If we were to retool the refactored script for another purpose, we would really only need to edit the arrays and use the search and replace function to update the array names in the loops.