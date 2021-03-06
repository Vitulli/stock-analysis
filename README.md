# Stock Analysis and Tool
##  Overview of Project
##### There is great interest for investment in today’s green energy market.  With that in mind the question was posed “What green energy stock show mass appeal for investment and strong returns?”.

##### To fulfill this request a tool was created to quickly shuffle through two years of green energy stocks for the years 2017 and 2018 for each day of active trading.  The data set had Ticker Symbol, Date of Trade, Opening Price, Closing Price, the Highs and Lows of the day, Adjusted Close, and lastly Volume Traded for the day.

## Results, Findings, and Comments
#####  To address the concerns of investment directly and generally, the year 2017 showed remarkable gains for almost all stocks.  TERP was the only company that had losses for the year at -7.2% whereas most other companies had double and triple digit gains.

##### 2017 Green Energy Gains and Losses
![](OrigVBAScriptFormatted2017.png)

#####  Despite the terrific gains demonstrated in 2017 almost all stocks tumbled severely in 2018.  Only ENPH and RUN had gains for 2018 but their respective gains were still fantastic.

##### 2018 Green Energy Gains and Losses
![](OrigVBAScriptFormatted2018.png)

#####  Overall, in recommending a stock for purchase the simple statement would be only to buy ENPH and RUN as they were the only two stocks that had positive gains for both years.  This is a very simple analysis and many of the stocks still retain a two-year gain over initial purchase.  Again, the safe bet would be the two stocks mentioned above.

##### To conduct this analysis a tool was created with two different methodologies and how these operate create large discrepancies in their execution times.  These will be discussed at length in the summary.  Just note that the speeds for execution of the refactored scripts eclipse that of the original script.

#### 2017 Original VBA Script
![](OrigVBAScriptAnalysisandTimetoRun.png)
#### 2018 Original VBA Script
![](OrigVBAScriptAnalysisandTimetoRun2018.png)
#### 2017 Refactored VBA Script
![](RefactVBAScriptAnalysisandTimetoRun2017.png)
#### 2018 Refactored VBA Script
![](RefactVBAScriptAnalysisandTimetoRun2018.png)

## Summary
##### As can be seen from the data sets above refactoring the code made a tremendous difference in the amount of run time for the script and while this may seem unimportant if the data set is expanded for more years or companies the run time could become quite large.  This is also significant because the timer in my original code is only set to run during the analysis loop whereas the timer counts the entire code set for the refactored version.

##### The refactored code demonstrates roughly a 6.5 to 8.5 times improvement over the original script for 2017 and 2018.  Again, if this code is run to a larger scale this will become important.

#####  To address the time savings the fundamental difference between these two pieces of code is the way it loops through the dataset, writes to excel, and formats.  In the first script, the code loops through the entire dataset looking for a predetermined sticker symbol and acts on the code's instructions during the loop.  The script then writes to excel cells, for that ticker symbol, individually or one by one.

##### The refactored counters this method by looping throught the dataset and storing the requested values in array.  The biggest difference is not writing values to excel until the end of the data collection code and then writing the stored values in their respective arrays.  It seems that writing to cell, as next to last instruction set, is the biggest time expenditure in the program and leaving this procedure to the end and after data collection creates huge time cycle time savings.

##### One last point of observation about these tools is that they both depend on a clean dataset to operate.  The tool is dependent on the sticker symbols being grouped together and in order.   Therefore, this code will only operate on a well managed dataset. 

##### Lastly, the original script has some of my own interpretations to it.  I used a button for each subroutine and just to experiment with the features of VBA and hence maybe different from the original request.
