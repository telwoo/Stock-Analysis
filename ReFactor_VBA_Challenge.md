# Stock Analysis VBA Excel

## Overview of Project

### Purpose
Steve wanted to outline specific stocks to make sure that his parents had the financial coverage they needed in order to enjoy their finances for their future. More specifically, there were 12 stocks from 2017 and 2018 to review with Steve's parent's to see if they were good to invest in. By refactoring (or editing) the previous format, we are able to make the VBA more quicker and outline the details in a more sufficient manner. The overall goal was to retrieve the ticker value (stock ID), outline the starting and closing price of the ticker, and review the return (based on year) of each stock.


## Results
After running the macro, it's determined that "ENPH" stock ticker has the highest daily volume of 607,473,500, with a positive return of 81.9%. "AY" stock ticker has the lowest daily volume of 83,079,900, with a negative return of 7.3%. Based on these outputs, it would be wise for Steve to detail to his parents that this **ENPH** would be a good stock for them to invest in in the long run. In 2017, there were more positive returns (11) from the stocks, and only one negative return - so it appears that the stocks were doing well that year! However, there were more negative returns from the stocks (10), and only two positive returns. We can conclude that there might been some large financial undertakings or economic influences that causes this large shift in return within the two years.

Below is an image of the original code before refactoring. One key item that I noticed is that within the original code, the format of the numbers within the "Total Daily Volume Column" was not as detailed as the refactoring. 
Before Refactoring Image:
<img width="767" alt="greenstock_formatting" src="https://user-images.githubusercontent.com/106715923/174691373-0aa157f9-5593-4a6d-a2ee-f4ed5e21b528.png">


After the refactoring, you can see that the numbers look more precise and are outlined to be more readable to Steve and his parents. In both original and refactoring images, the initial input text box that allows in the macro to select the year allows the data to be specified. This one detail outlines the profitability of the stock from one year to another. 


After Refactoring Image:
<img width="799" alt="VBA challenge_formatting" src="https://user-images.githubusercontent.com/106715923/174691419-41433200-dcf5-4e3f-ac1e-601e4f5d56b4.png">

 
## Summary

- What are the advantages or disadvantages of refactoring code?
 1. Advantage(s): formatting of the arrays created for tickers, tickerVolumnes, tickerStartingPrice, and tikcerEndingprice. This make the numbers of the "Total Daily Volume" column look more fluid and readable. 

 2. Disadvantage(s): the run time for the macro has decreased significantly. The time it takes to run appears be SLOWER than the original scripting. Will a few more modifications to the formatting or changes to my computer's processing, I'd believe there could be a way to make the lapse time quicker. See screenshots below:
Original code lapse time for 2017:

<img width="263" alt="green_stock_2017 lapse time" src="https://user-images.githubusercontent.com/106715923/174691502-26eb51f8-ec18-4860-9cbd-b1db7bb0facf.png">


Refactoring code lapse time for 2017:

[VBA_Challenge_2017 elapsed run time](https://user-images.githubusercontent.com/106715923/174691522-ef0c4559-f5bf-41d5-87dd-6cc93c1ede84.png)


- How do these pros and cons apply to refactoring the original VBA script?
 1. Pros: Refactoring helps in finding bugs within the system; refactoring can help make programming faster and easier to understand/maneuver through. Because I was able to outline the formatting so that it was be more readable to the viewers, it helps in reviewing it better than the original VBA script.
 2. Cons: sometimes in debugging, you could easily spend copious amounts of time refactoring; in a corporate sense, it could be the possibility of running out of resources to fund refactoring. It took a longer amount of time to add a few more codes to the refactoring script than the original script (which could be cumbersome after already trying to make sure that the code is working properly in the first place). 
