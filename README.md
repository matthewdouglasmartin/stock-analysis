# Stock Analysis with VBA

## Overview of Project

VBA in Excel can be used to automate Excel functions and formulas in an effort drastically minimuze labor intensive formatting and the inputting of data.  Steve is helping his parents analyze the stock market so that they can make investment decisions which will result in positive returns.  Initially, his parents were interested in one specific investment opportunity with DAQO New Energy Corp (DQ), a company that resonated with them due to many of their green energy initiatives.  After initial review of DQ's stock performance in 2018, the results showed a -62.6% return.  

This then prompted Steve and his parents to expand the analysis to include at least a dozen more companies, potentially more in the future, seeking a more profitiable investment opportunity.  The initial VBA code and Excel spreadsheet provided to Steve and his parents that provided a full analysis of 12 company stocks was functional but its efficiency was not optimized for this dataset and especially not for a larger dataset that might be used in the future.  Below are the results of the refactored VBA code provided to Steve and his parents, which as will be shown, out performed the initial VBA code. 

## Results

The analysis of stocks for Steve and his parents was conducted on a dataset with the results from fiscal 2017 and 2018.  Starting with the results from 2017, about 92% of the companies analyzed saw positive returns at the end of 2017.  Solely looking at this extracted data, though showing great performance, does not provide insight into company and/or market trend.  The stock for these companies performed very differently during 2018, only 2 of the 12 companies analyzed saw a positive return on their stock value.

Initially, the VBA code used to analyze the stock performance in 2017 and 2018 ran smoothly, but not very fast, and it did not format the compiled results in a visually accesible manner.  Here are two screenshots of the first attempt to run the VBA code for both 2017 and 2018:

2017
![Initial 2017 Stock Analysis with VBA Results](https://github.com/matthewdouglasmartin/stock-analysis/blob/master/Resources/VBA_Initial_2017.png)

2018
![Initial 2018 Stock Analysis with VBA Results](https://github.com/matthewdouglasmartin/stock-analysis/blob/master/Resources/VBA_Initial_2018.png)

As can be seen from these images, the runtime of the analysis is not extremely slow, it can still be performed in less than 1 second.  If more companies' stock performance are added in the future, the time to run the VBA code will only increase.  In addition, it lacks formatting that would enhance the ability to quickly glean valuable performance statistics.

After refactoring the initial VBA code, the runtime decreased from about .699 seconds (2017) and .734 seconds (2018) to .152 seconds (2017) and .160 seconds (2018).  The decrease in runtime occured after successfully refactoring the code and adding formatting code within the sub routine to output the results in a more "readible" and visual format.  Beleow are two screenshots of the results after running the refactored VBA code that can be compared to previous screenshots referenced:

2017 Refactored
![Refactored 2017 Stock Analysis with VBA Results](https://github.com/matthewdouglasmartin/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

2018 Refactored
![Refactored 2018 Stock Analysis with VBA Results](https://github.com/matthewdouglasmartin/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

Once the VBA code had been refactored to run faster and output the analysis in a user friendly format, comparing the performance of all the stocks during 2017 and 2018 is much less tedious.  As mentioned above, about 92% of the companies saw positive returns at the end of 2017 while only about 17% of the companies saw positive returns in 2018.  This is a big shift in the market/company performance and sheds some light on how companies can be expected to trend in the future.  Though there are a lot of companies that had stock which performed well in 2017, there are only 2 companies that saw positive returns in both 2017 and 2018, ENPH and RUN.  Even though ENPH's stock performed well in both 2017 and 2018, its return dropped by about 38%.  On the other hand, RUN's stock had an increase on its positive return from 2017, increasing from a 5.5% return in 2017 to 84.0% return in 2018.  Ideally more years of company stock performance should be utilized to better identify positive trends.  Relying only upon the information provided, guiding Steve and his parents to invest soley in RUN is the best conclusion that can be drawn from the results provided by the analysis.  

## Summary

As described above, VBA code (and any code used in programming for that matter), can benefit from the process of refactoring.  The end result is often faster, easier to understand, and more flexible.  Though there are often more advantages to refactoring code, the disadvantages of refactoring sometimes carry more weight.  Most often the main reasons for not investing in refactoring code are the result of lacking in money and time.  These two disadvantages weigh heavily on the teams and companies in charge of refactoring code.  Below are some examples of the advantages and disadvantages of the VBA code before and after refactoring. 

Refactoring the intial VBA code to address concerns regarding run time and formatting, makes the VBA code significantly more efficient, flexible, and user friendly.  Addressing the concerns around effciency and flexibilty of the VBA code can primarily be resolved by using multiple arrays and an index tracker.  Referencing these arrays and an index tracker in the "for" loop, along with "If-Then" conditionals within the "for" loop, allow the refactored VBA code to process all the rows of stock performance data very quickly.  Once all of the requested data is pulled from the 2017 dataset with this update, an additional "for" loop is added which loops through the arrays of pulled data and outputs it to the "All Stock Analysis" worksheet.  

In contrast, the initial VBA code only used one array, the "tickers" array, and no index.  When the "for" loop is inititiated it assigns the data to one array and that one array is responsible for holding and outputting all of the sought after results.  Though this VBA code works and performs the task needed, organizing the pulled data within multiple arrays speeds up the VBA code's ability to reference the same amount of data more quickly.  Not only does the efficiency benefit from this update to the VBA code, but so does its flexibility, more stock performance data could be added to the dataset in the future with minimal negative impact to the run time.

Since the runtime of the VBA code has decreased, addtional VBA code can also be added to this subroutine to aid in the formatting of the results in a more user friendly manner.  Doing so tidys up the results output in the new worksheet.  The newly formatted results can now be accessed by a wider audience of users, including individuals that do not have much experience in stock performance analysis.

 These advantages gained are well worth the work put into refactoring the VBA code.  As noted above the main disadvantage to refactoring the VBA code was the time it took to do so successfully.  There was a great deal of trial and error and collaboration amongst team members to land on the final working code.  Had the required amount of time or the collaborative efforts not been available, refactoring the VBA code might not have been deemed truly necessary or a viable expense.  
