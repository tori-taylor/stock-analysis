# Stock Analysis - Module 2

## Overview of Project

### Purpose
The purpose of this analysis is to help Steve do research on stocks, which includes their yearly reutrn and total daily volume. In this task specifically, the code was edited from the previous code to make it work faster, therefore Steve would be able to expand his research to more stocks and still be able to run the code in a small amount of time. 

## Results

### Run Times
The refractored code (the edited code) worked faster than the original code. The orignal code for 2017 and 2018 ran in 0.543 s and 0.453 s respectively. While the refractored code for 2017 and 2018 ran in 0.066 s and 0.070 s respectively. While using a small number of stocks this wouldn't appear as much of a difference but when looking at a significantly more stocks this time difference would start to add up! In addition the new code includes formatting, while the orignal code does not. These images show the run times of each code. In order: Origial 2017, Original 2018, Refractored 2018, Refractored 2018. 

![VBA_Challenge_2017.original](VBA_Challenge_2017_original.PNG) ![VBA_Challenge_2018original](VBA_Challenge_2018_original.PNG) ![VBA_Challenge_2017](VBA_Challenge_2017.PNG) ![VBA_Challenge_2018](VBA_Challenge_2018.PNG) 

### Advantages and Disadvantages of Refractored Code
The advantages of refractored code is that it is more efficent, has a faster run time, uses less memory and is easier for another person to interpret. The disadvantages are it is time consuming to edit your code, and is sometimes more complicated.

### Advantages and Disadvantages of this analysis original code and refractored code.
The main disadvantage for me with this refractored code is that I had trouble getting it to work, I used the code found here https://github.com/caseychen3605/stock-analysis to help me when I got stuck. The issue I had was the following:
'2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i
Here, I didn't orignally understand what initializing all the tickerVolumes to zero meant. The overall advantage of this code is that it goes through the data once, while the original code went through all the data twelve times, one time for each ticker. Even though refractored code is said to usually be easier to understand, I found this refractored code more difficult to understand than the original.
