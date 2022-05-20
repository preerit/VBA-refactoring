# VBA-refactoring

## Overview of Project: Explain the purpose of this analysis
The overall purpose of this analysis is to analyse the performance of different stocks and to automate the calculation of performance parameters such as total volumes, starting and ending prices. The automation is done using VBA Macros. The reason for automation is to avoid repetition on 2 fronts -
* Repeating excel analysis for different stocks
* Repeating the same for another year

In this activity in particular, we are refactoring the code, meaning, we are making the code more efficient and letting it run faster than it did before. We are measuring code performance (run time) after coding, that lets us understand if it is refactored efficiently.

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
### On stock performance:

Stock performace in 2017 is as follows:

![StockPerformance_2017](https://github.com/preerit/VBA-refactoring/blob/main/StockPerformance_2017.png)

Stock performace in 2018 is as follows:

![StockPerformance_2018](https://github.com/preerit/VBA-refactoring/blob/main/StockPerformance_2018.png)

* 2017 has higher returns than 2018 for all stocks in general
* However, the total daily volume in 2018 is higher than that in 2017

### On execution times:
The original code for 2017 has the run time as follows:

![OGcode_2017_runtime](https://github.com/preerit/VBA-refactoring/blob/main/OGcode_2017_runtime.png)

The refactored code for 2017 has the run time as follows:

![VBA_Challenge_2017](https://github.com/preerit/VBA-refactoring/blob/main/VBA_Challenge_2017.png)

The original code for 2017 has the run time as follows:

![OGcode_2018_runtime](https://github.com/preerit/VBA-refactoring/blob/main/OGcode_2018_runtime.png)

The refactored code for 2017 has the run time as follows:

![VBA_Challenge_2018](https://github.com/preerit/VBA-refactoring/blob/main/VBA_Challenge_2018.png)

* The refactored code has a more efficient run time for both 2017 and 2018.
