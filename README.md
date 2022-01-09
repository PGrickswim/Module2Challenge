# Analysis Of Stocks
By Richard Sadowski

## Overview of Project
By using Visual Basic for Applications (VBA) to work with large amounts of data, this report seeks to draw conclusions about the performance of a group of "green" stocks through the years 2017 and 2018. 

## Purpose
The purpose of the project was to refactor code written specifically to "do the job" of analyzing a group of 12 stocks, and make it ready to analyze a larger data set without using more computing power than necessary and keeping the program as simple as possible when run.

-By including a timer in the code, it is easy to analyze the amount of time it took to run
-By analyzing the output of the code, it is easy to spot if there are any differences in the output caused by refactoring.
-Additionally, looking at the code itself can provide a view of the complexity of the original and refactored code.

## Results

### Original Code

- The original code used to analyze stocks leveraged a nested for loop to first view the entirety of rows in the dataset and then for each row, find the appropriate starting or ending price based on the layout of the data.
- The code continued to loop until all rows were exhausted based on how the original code was written.

**Original Code Script**

![Original Code Script](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Origina_Code_Script.png)

### Refactored Code

- The refactored code used to analyze stocks leverages multiple for loops but they are not nested in each other, meaning the computing power needed to scan each array is less than that of the original code.
-Once the refactored code finds data for one of the defined tickers, the tickerIndex is increased by one so that the code can then look to find the next ticker.

**Refactored Code Script**

![Refactored Code Script](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Refactored_Code_Script.png)


### Outputs

- Below, see that the outputs, after programming and formatting code are run, are identical in output.

**2017 - Original Code**

![2017 - Original Code](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Stocks_2017.png)

**2017 - Refactored Code**

![2017 - Refactored Code](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Stocks_2017.png)


**2018 - Original Code**

![2018 - Original Code](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Stocks_2018.png)

**2018 - Refactored Code**

![2018 - Refactored Code](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Stocks_2018.png)


### Execution Times
- See below that while execution times, as measured by the timer written into the VBA Script, differ only slightly, there is a time advantage to running the code with the refactored script. 

**2017 - Original Code**

![2017 - Original Code Time](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Original_Code_2017.png)

**2017 - Refactored Code**

![2017 - Refactored Code Time](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/VBA_Challenge_2017.png)

**2018 - Original Code**

![2018 - Original Code Time](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Original_Code_2018.png)

**2018 - Refactored Code**

![2018 - Refactored Code Time](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
In summary, the refactored code sets us up well to analyze a larger dataset. Refactoring code has some advantages and disadvantages, which will be shown below.
1. Advantages
    - Can allow a program to run more quickly
    - Can build a program to be able to be modified more easily in the future
    - Can make the programming more easy for other programmers to understand
    - Can adjust formatting of code to meet agreed-upon standards
2. Disadvantages
    - If code was working before, refactoring introduces changes that may require additional troubleshooting
    - Refactoring takes time, and does not immediately change the output of the program
As with all code writing, captioning out comments and being mindful of how white space is used is paramount to establishing a coherent code that can be understood by other programmers.

## Conclusion
As it relates to this application, the refactoring of the code did not have a significant impact on how VBA handles the current dataset. If the analysis were to end here, the original code would be sufficient. However, with an eye toward applying additional data so that the entire stock market can be analyzed, this adjustment might make the code run faster for this other application. Still, since the variable "tickers" needs to have each stock ticker initialized to a number, doing this for a larger dataset containing hundreds or thousands of stocks would be difficult to run this way; the setup of the code would be extreme. Also, based on how this code works, it is important that the data is sorted by ticker before the program is run. It would be helpful to have code that can look at stocks over a large amount of time when the data is not sorted in this way.




