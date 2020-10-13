# Stock-analysis
Module #2 

# Overview of Project

Earlier for our client, Steve, we prepared an Excel workbook using VBA code that would help Steve analyze the performance of various “green energy” stocks for the years 2017 and 2018. The purpose of this was because Steve’s parents had invested heavily in one “green enrergy”  stock, “DAQO,” and wanted to know if this was a wise investment, and if not, what would be a good alternative “green energy” stock that the could potentially invest in instead. 

Steve was very happy with our analysis, but now to do a little more research for his parents, he wants to expand the dataset to include the entire stock market for the years 2017 and 2018. 

To do what Steve is asking is the basic premise of this project. In order to do what Steve is asking, our project focuses on:
	- Refactoring (editing) our original VBA code to loop through all the data one time to collect the following information for 		each stock (Ticker name, Totally Daily Value, percentage of Return) for both 2017 and 2018, as well as the time it took to 		do so using our code. 
	- Analyzing our findings after refactoring. 
We will then summarize the advantages and disadvantages of refactoring code, and how it applies to the refactoring our original VBA script.

# Results

After refactoring our code and running it to analyze our data, our results were as follows:

For 2017 
- Insert VBA_Challenge_Stock_Analysis_2017

Our time for running our code
- Insert 2017 time

For 2018  
- Insert VBA_Challenge_Stock_Analysis_2018

Our time for running our code
- Insert 2018 time

Viewing our data, we can determine that the following changes in Total Daily Value happened between the end of 2017 and the end of 2018:

Total Daily Value decreased: 
- “AY” 
- “CSIQ”
- “FSLR”
- “JKS”
- “SPWR” 
Total Daily Value increased 
- “DQ”
- “ENPH”
- “HASI”
- “RUN”
- “SEDG”
- “TERP”
- “VSLR”


However, we can also see that from the end of 2017 to the end of 2018, only 2  stocks, “ENPH” and “RUN” yielded a positive return percentage. 

We can also see that the stock, “TERP” had an increase in its return percentage in 2018  compared to in 2017, but still yielded a negative return percentage. All our other stocks had fairly dramatic drops in its return percentages. 

Based on these findings, one might argue that the stocks “ENPH” and “RUN” are likely stocks that Steve should consider investing in. However, it can be argued that we could use more information. For example, we have no way of knowing based solely on the data why exactly there was such a dramatic change in Return Percentage from 2017 to 2018 that can be attributed to other forces, and if so, what were they? Are they quantifiable? 

Unfortunately, we did not record the execution time our script took before refactoring, so we have no comparison to give in regards to how much faster or how much slower the times are after refactoring. Apologies. 

# Summary

After our refactoring our code, we can see that the code does work, and Steve can make some good general assumptions based on its output, but would benefit from asking more questions and finding out their answers in regards to real world events that may have contributed to the outcomes in question. 

In regards to refactoring code, we can argue the following:

Advantages to refactoring code are that we can save ourselves time, as our pre-existing code lays a good ground work to start with. 

Disadvantages to refactoring code are that if you were not the original author of the code, and the original author didn’t do a good job of annotating what each segment of code is for, it could be confusing, and negate the time saving aspect. The original code itself could also be incorrect, causing one to lose even more time troubleshooting.

This relates to our own refactoring experience, as the original VBA code prior to beginning this project was incorrect, and there was a large portion of time spent fixing its errors. By the time all the errors were corrected, it was not beneficial or easy to go back to the original code to get its run times to compare to the run times with the refactored code. 
