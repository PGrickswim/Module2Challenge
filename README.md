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
![Original Code Script](https://github.com/PGrickswim/Module2Challenge/blob/main/Resources/Origina_Code_Script.png)