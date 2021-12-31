# VBA Challenge

## Overview of Project

The purpose of this project was to gain experience with VBA while learning about ways to make code more efficient and expedient. This involved working through the creation of various macros within VBA, then refactoring and improving those macros to decrease runtime.

## Results

It is clear that the execution time of the analysis dropped dramatically after refactoring. The original runtimes are shown below:

![2017original](resources/2017_original.png)
![2018original](resources/2018_original.png)

while the runtimes for the refactored code are shown below:

![2017refactor](resources/VBA_Challenge_2017.png)
![2018refactor](resources/VBA_Challenge_2018.png)

The execution time for the refactored script was roughly 10% of the execution time for the original code. The difference between 0.6 and 0.06 seconds might not be very noticeable, but changes of this magnitude are extremely noticeable when applied to longer, more complicated scripts. For example, if a script's runtime can be decreased from 10 minutes to 1 minute, that is a huge improvement.

This change in runtime is largely due to the use of arrays instead of a nested for loop. When the following code is used to create arrays for volume, starting price, and ending price

    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
then there is no need for a nested for loop to run through each ticker, such as the one below:

    for each ticker
        for each row
        next row
    next ticker
    
This speeds up the code dramatically, as for loops tend to increase runtime.
