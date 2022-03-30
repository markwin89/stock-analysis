# Stock Analysis

## Overview

### Purpose
Analyzing the return percentage and the daily volume of the stock tickers under the 2 year parameters.

## Results
### 2017
2017 had great returns with the exception of TERP.  The refactored code ran for 0.1015625 seconds to execute. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/101272613/160852044-0feef5cf-11f8-4000-a934-66c11494bbb8.png)

### 2018
2018 was a bad year for most stock tickers except ENPH and RUN.  The refactored code ran for 0.0859375 seconds to execute. 
![VBA_Challenge_2018](https://user-images.githubusercontent.com/101272613/160852061-163366f3-b323-4882-8507-abf6023c2088.png)

---
## Advantages and Disadvantages of Refactoring Code
### Advantages
-Refreshes the code making it a bit easier to understand, consume and maintain. 
	-Provides a cleaner design.

### Disadvantages
-Time consuming if the orginal script was long and detailed. 
-Risky due to bugs.  The orginal script could have bypass a bug but the refactor code caused it to bug again. 


### Advantages of the original/refactored VBA script
-It took less time to execute the refactored script then the original.  
-It was able to pull a msgbox to enter the year for the ticker dataset instead of writing a longer code to list both worksheets. This is also cleaner to view since there is only one tab to review "All Stock Analysis". 


### Disadvantages of the original/refactored VBA script
The code had to be written where it referenced the correct array or it will pull a bug Overflow error. 


  

