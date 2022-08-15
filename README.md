# VBA of Wallstreet

## Overview
    This project involves the creation of a VBA script designed to collect multiple peices of data from a spreadsheet containing stock market data for multiple public companies, and output that data in a readable format to another worksheet for analysis.

### Purpose
    The scripts purpose is to calculate the total yearly volume and the yearly return for each individual company in a stock market dataset. The data is then arranged in an output worksheet by each companies ticker, and formatted to make the analysis easier. The client intends to use the script to quickly analyze a large number of public companies for potential investment opportunities. The original script written was capable of adequately handling a dataset containing a small number of companies, but the customer desires to implement it for a much larger dataset. The script needed to be refactored in order to efficiently collect the data from larger datasets. Here we see the results of that refactoring. 



## Refactoring the VBA script
    The original script used nested `For` loops and a single `array` of tickers to pass through the dateset to calculate the yearly volume and return, and output that data to the worksheet one ticker at a time. The orignal script did allow for the user to select the year of data on which to run the analysis, but did not format the output worksheet. While this was sufficient for a small dataset, larger datasets analyzed with this method would have been very time consuming. 
    
    The code was refactored by creating an `index` variable and three additional arrays to hold the necessary outputs for each ticker in the dataset. A `For` loop was then written to pass through the dataset one time, collecting and storing all the data for each ticker along the way using the `index` variable to store the outputs in the three output arrays. Then another `For` loop was created to output the data from the three arrays, along with the tickers from the ticker array, to the analysis worksheet. The script also formats the worksheet into a easily readable format, and allows the user to select which year of data to run the analysis on.   

### Results
    The orignal script needed to loop through the dataset an additional time for each new ticker's data to be collected. This is a very time consuming method. A dataset of only 12 companies took longer than half a second to be analyzed. An increase in dataset size would have led to an exponential increase in the time required to analyze it. A dataset that contained thousands analyzed with this method would have been very inefficient. The below time stamps are the original script run on two years of data for 12 companies. 
'Input original script time stamps'

    Once refactored, the script only needs to pass through the dataset once to collect all the necessary data. Then to pass through the output arrays to output the data to the worksheet. This results in a much faster runtime, and an increase in dataset size will not have such a drastic effect on the time needed to run the script. The refactored script is also easier to understand, simpler to update and maintain. Here we see the drastic differences in runtime for the same dataset analyzed with the refactored script. 
'Input refactored script time stamps'

## Summary
    Refactoring code has some great advantages. It can result in code that is easier to read and understand. It can result in code that is much more efficient to run, requiring less memory or processing power, or taking less time. It can also make the code easier to update and maintain in the future. 
    At the same time refactoring comes with some disadvantages. Refactoring requires a complete understanding of what the original script is doing, and why it is doing it. Refactoring can be very time consuming, and has the potential to produce more bugs to the code or even break the code altogether.
    As the original VBA script for this project was refactored, these came into play. Understanding what the original code was doing, and trying to create the same results in a more efficient manner was challenging. Once the new method was implemented, it had several bugs that caused it to fail. Some of these bugs were the result of reused code that had not been rewritten appropriately, others were the result of new additions poorly implemented. However, once these challenges were overcome the results were significant. The code is much easier to understand from an outsiders point of view, and much easier to maintain and update to handle new datasets. Most importantly the new method employeed runs much more effeciently than its predecessor, resulting in significantly decreased run times for the script. The refactored code will be able to adapt to many different datasets easily, and handle datasets of varying size without significant increases in runtime.  
