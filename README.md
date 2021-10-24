# Stock Analysis

## Overview of Project 
The purpose of this project is to compare performance across twelve different stocks across 2017 and 2018. This includes examining the Total Daily Volume, the return of each stock over the year and comparing returns from 2017 and 2018.   
This analysis was done using Visual Basic for Applications (VBA), an excel based programming language that can be used to help automate analysis. In addition, to examining the returns of different stocks, a comparison will be done between original VBA script and refactored VBA script to determine which one ran faster. 

## Results
### Stock Performance
Overall, most of the stocks in this analysis performed significantly better in 2017 than in 2018. As can be seen in the below screenshots only two stocks performed better in 2018 than 2017, RUN and TERP. However, there are only two stocks that provided positive returns in 2018, ENPH and RUN, with a 81.9% and 84% return respectively. When compared against the Total Daily Volume, both ENPH and RUN had significant increases in the volume of stock being traded. ENPH's Total Daily Volume in 2018  by 174% while RUN's Total Daily Volume increased by 88%, when compared to 2017. This large increase in Total Daily Volume can indicate that there is an increased level of interest in the respective stock, especially by large institutional investors. This interest can help drive up the price of the stock could indicate that it is a worthwhile stock to invest in. However, the Total Daily Volume should not be considered as a sole factor. DQ also had a large increase in Total Daily Volume between 2017 and 2018 but its return is down significantly. Given that this dataset is looking at a full year it is possible the data is slightly skewed. DQ performed much better at the beginning of 2018 and its performance did not start declining until later in the year. Going forward it might make more sense to look at stock performance trending by month or quarter to be able to identify changes in performance faster. 


![2017 Screenshot](https://user-images.githubusercontent.com/91712554/138568317-e222b64b-317e-458a-ae36-778161dd755a.png)        ![2018 Screenshot](https://user-images.githubusercontent.com/91712554/138568320-db2b4f4e-f605-4d42-8d62-c6679b0d128a.png)

### Execution Times
This analysis was completed two times, first using an original set of script, and then using refactored script. Refactoring script includes revising and restructuring existing code with the intent to improve the overall functionality. In this case the original code included a nested For statement, meaning the computer has to run each line of code for the number of loops indicated by the variables. In the refactored version the code is set up to only have to loop once in order to collect and output the same data as the original. The screen shots below show the comparison of the nested For statement and the refactored code which uses the variable tickerIndex. 

InputScreen Shots

The refactored code ran consistently faster than the original script. The original script took 0.71 seconds for 2017 and 0.68 seconds for 2018. While the refactored code took 0.11 seconds for 2017 and  0.30 seconds for 2018. You can see the screen shots below that show the MsgBox output indicating the time it took to gather all date and create the outputs.

Screenshots!!

## Summary 
### Advantages and Disadvantages of Refactoring code
The advantages of refactoring code include that it can help improve the overall design and format of code, potentially making it easier to understand. It can also help discover bugs in the code and improve the run time. 

### Potential disadvantages of refactoring include that it can be a time consuming process, especially if the application is large. This could take away from other, more valuable, uses of time. There is also the possibility that the individual refactoring does not understand the overall goal of the code. Refactoring the code in sections could have unintended consequences for later parts of the code. 

### Pros and Cons of Refactoring the Original VBA Script
Refactoring as it relates to this specific code definitely ultimately improved the overall run time of the code. However, given that in this project we are analyzing a relatively small amount of data the original script did not take long to run. Unless in the future this code was going to be used on a larger scale or as a part of a larger application the time take to refactor and debug the code might not have been worth the marginal increase in run time.  


