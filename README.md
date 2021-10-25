# Stock Analysis

## Overview of Project 
The purpose of this project is to compare performance across twelve different stocks from 2017 to 2018. This includes examining the Total Daily Volume, the Return of each stock and comparing results between the two years.  
This analysis was done using Visual Basic for Applications (VBA), an excel based programming language that can be used to help automate analysis. In addition to examining the returns of different stocks, a comparison will be done between original VBA script and refactored VBA script to determine which one had a faster run time and why. 

## Results
### Stock Performance
Overall, most of the stocks in this analysis performed significantly better in 2017 than in 2018. As can be seen in the below screenshots only two stocks performed better in 2018 than 2017, RUN and TERP. In addition, there are only two stocks that provided positive returns in 2018, ENPH and RUN, with a 81.9% and 84% return rate respectively. When compared against the Total Daily Volume, both ENPH and RUN had significant increases in the volume of stock being traded. ENPH's Total Daily Volume in 2018 increased by 174% while RUN's Total Daily Volume increased by 88%, when compared to 2017. A large increase in Total Daily Volume can indicate that there is an increased level of interest in the respective stock, especially by large institutional investors. This interest can help drive up the price of the stock and could indicate that it is a worthwhile stock to invest in. However, the Total Daily Volume should not be considered as an independent factor when considering which stocks to invest in. DQ also had a large increase in Total Daily Volume between 2017 and 2018 but its year over year return saw a sharp decline. Given that this dataset is looking at a full year it is possible the data is slightly skewed. DQ performed much better at the beginning of 2018 and its performance did not start declining until later in the year. Going forward it might make more sense to look at stock performance trending by month or quarter to be able to identify changes in performance faster and adjust an investing strategy appropriately. 


![2017 Screenshot](https://user-images.githubusercontent.com/91712554/138568317-e222b64b-317e-458a-ae36-778161dd755a.png)        ![2018 Screenshot](https://user-images.githubusercontent.com/91712554/138568320-db2b4f4e-f605-4d42-8d62-c6679b0d128a.png)

### Execution Times
This analysis was completed two times, first using an original set of script, and then using refactored script. Refactoring code includes revising and restructuring existing code with the intent to improve the overall functionality. In this case, the original code included a nested For statement, meaning the computer has to run each line of code for the number of loops indicated by the variables. In the refactored version the code is set up to only have to loop once in order to collect and output the same data as the original. The screen shots below show the comparison of the nested For statement and the refactored code which uses the variable tickerIndex. 

![Original Script](https://user-images.githubusercontent.com/91712554/138604443-e936d927-4ad1-471e-9666-45e4a930b72e.png)   ![Refactored Script](https://user-images.githubusercontent.com/91712554/138604445-c592eb5d-a647-4319-8afe-7857203896e2.png)

The refactored code ran consistently faster than the original script. The original script took 0.71 seconds for 2017 and 0.68 seconds for 2018. While the refactored code took 0.11 seconds for 2017 and  0.30 seconds for 2018. You can see the screen shots below that show the MsgBox output indicating the time it took to gather all data and create the outputs.

![Timer 2017](https://user-images.githubusercontent.com/91712554/138604452-17d591a5-e5d1-4629-af5c-b7ef41ed7fe0.png)  ![VBA_Challenge_2017](https://user-images.githubusercontent.com/91712554/138604466-7db5d495-49f1-4a40-9a25-77046df8eae7.png)



![Timer 2018](https://user-images.githubusercontent.com/91712554/138604457-19077bf0-c1b3-488e-a90d-ed79ba22c29f.png)   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/91712554/138604472-17234def-f7d7-40c2-9e31-69ba7f33a174.png)

## Summary 
### Advantages and Disadvantages of Refactoring code
The advantages of refactoring code include that it can help improve the overall design and format of code, potentially making it easier to understand. It can also help discover bugs in the code and improve the run time. 

Potential disadvantages of refactoring include that it can be a time consuming process, especially if there is a lot of code for a larger application. This could take away from other, more valuable, uses of time. There is also the possibility that the individual refactoring does not understand the overall goal of the code. Refactoring the code in sections could have unintended consequences for later parts of the code. 

### Pros and Cons of Refactoring the Original VBA Script
Refactoring the code for this specific project ultimately improved the overall run time of the code. However, given that in this project we are analyzing a relatively small amount of data the original script did not take long to run. Unless this code was going to be used on a larger scale, or as a part of a larger application in the future, the time needed to refactor and debug the code might not have been worth the marginal increase in run time.  


