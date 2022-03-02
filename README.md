# All Stocks Analysis Refactored

## Overview of Project

We are creating an all stocks analysis for Steve and his parents for the years 2017 and 2018 to help them decide for their investment.  They can see the results at the click of a button after choosing a year.  We started our code by creating arrays for all stocks and after completion, it worked first time, but it took about .90 seconds to run. Then we refactored the code to run at a quicker speed by using arrays to generate results. Refactoring the code is to make the code more efficient - by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.  

## Results

We were successful in refactoring to reduce the run time to under .20 seconds. The original code's run time was around .90 seconds for each year as shown in images below:

!![AllStocks Analysis Run Time for 2017](https://user-images.githubusercontent.com/98566486/156426400-384a9cdc-5fdf-40cb-8362-a0583720c131.png)
!![All Stocks Run Time for 2018](https://user-images.githubusercontent.com/98566486/156426425-65803521-3724-45b3-ad30-8eb1f469ecf3.png)

We created the ticker index variable to use in new arrays that we created for this code.  We created three output arrays: ticker Volumes, ticker Starting Prices, and ticker Ending Prices. Then we used the nested For loops with conditional formatting to loop through the data to generate the final analysis for all stocks as shown in the code image below:

![Screen Shot 2022-03-02 at 1 57 11 PM](https://user-images.githubusercontent.com/98566486/156429541-6bda16c1-5936-4e27-afe2-6de06d5f6870.png)

Our analysis shows better overall stock returns in 2017 compared to 2018 all stocks return columns:

![Screen Shot 2022-03-02 at 1 52 28 PM](https://user-images.githubusercontent.com/98566486/156428783-f1ddbbb4-89bd-4bfe-a062-7893ccf81c72.png)
![Screen Shot 2022-03-02 at 1 52 08 PM](https://user-images.githubusercontent.com/98566486/156428824-86269154-c6d9-428f-8f14-a9cf87e0f75f.png)

Our refactored analysis code run time is under .20 seconds :

![VBA_Challenge_2017](https://user-images.githubusercontent.com/98566486/156429662-ca9bd3f8-f12a-433f-bb7e-19d9341c3c29.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/98566486/156429704-2d5c0fd2-0d7d-474a-bb48-1399aff037ea.png)

## Summary

There are some advantages as well as disadvantages of refacting a code in VBA. 
###  Advantages
The advantages are use of less memory, time-efficient by taking fewer steps, and improve the logic of the code to make it easier for future users to read. 

### Disadvantages
Some disadvantages are extra time and budget for developers, can be risky when developers do not understand the original code and cause interruption in original code. If the data is large refactoring will be risky. It may introduce bugs.

In our case, the data was not too large and refactoring the original VBA script resulted in a more efficient run time.  However, time spent in refactoring the code is quite a lot compared to run time results.  In this case, refactoring the code was not worth the time spent unless the stock data grows and needs to run hundreds or thousands of times in a short period. 



