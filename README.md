# stocks_analysis

### Overview of Project
This project was created to analyze the `Daily volume` and `return` for various green energy stocks in 2017 and 2018.  

### Results

#### Stock Performance for 2017 and 2018
If we compare the Return for the green stocks in 2017 and 2018, we can see that almost all stocks took a downturn in 2018.  2017 was an amazing year for many of the stocks with half of them returning 50% or greater returns and only one `TERP` showing a negative return
![Analysis of stocks 2017](/Resources/Stock_Analysis_2017.png)


However, 2018 had almost all of these stocks showing a negative return.  The exception to this is `ENPH` and `RUN` with both showing a positive return of more than 80% in 2018.
![Analysis of stocks 2018](/Resources/Stock_Analysis_2018.png)

#### Script execution times for original vs refactored
Refactoring the VBA script greatly reduced the script execution time for this macro.  Both years went from more than 1 second in the original script to less than a quarter of a second in the refactored script. 
![execution time for refactored script 2017](/Resources/VBA_Challenge_2017.png)


![execution time for refactored script 2018](/Resources/VBA_Challenge_2018.png)


### Summary

#### Advantages of refactoring code
As this example shows refactoring the code can greatly reduce the amount of time a code takes to run and the amount of space required to execute the code.  While this may not be significant in this relatively small file, it could have significant impact on a larger data set or time sensitive analysis.

#### Disadvantages of refactoring code
While refactoring code can reduce the time spent in analysis, it is important to weigh the amount of time it takes to do the refactoring.  In some cases, the time spent working the VBA code is not worth the results.

#### Applying to refactoring the original VBA script
In reference to refactoring our script `VBA_Challenge` it was certainly a benefit to reiterate the concepts learned in the module and to see the way that changes to the module effected the execution time. However, if this had been real life, I do not know that the changes made were of enough benefit in the small data set to warrant the time spent by the analyst.  

