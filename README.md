# Green Stock Analysis
Project to analyze and compare green energy stocks.

## Overview of Project
### Purpose
I conducted a green energy stock analysis through excel and excel's VBA, in order to help a friend named Steve who just graduated with a degree in finance. Steve's first clients will be his parents, who are passionate about cleantech companies, specifically those centered around renewable energy. Steve's parents put a lot of their investment money into one company's stock without doing much research, so Steve wants to ensure his parents diversify their portfolio to maximize their return. However, Steve is not familiar with VBA, so I took the 2017 and 2018 green energy stock data Steve pulled and automated some analyses around stock trade volume and stock return to help Steve in his portfolio analysis and reduce possible errors that manual calculations can cause.

## Results
### Stock Performance
The analysis looked at two different years for the selected green stocks, 2017 and 2018. The first picture below is the stock volume and return analysis for 2017. This year seems to have been an exceptional year for the green energy industry because 93% of stocks had positive returns. In fact, 27% of stocks had over 100% returns. The stock that Steve's parent's invested their money into has the ticker DQ, which had the highest return of all the selected stocks at almost 200%.

![2017_stock_performance](https://github.com/asliwinski23/Stock-Analysis/blob/main/2017_stock_performance.png)

The story for the 2018 analysis is not so bright, as seen in the image below. Only 13% of stocks had a positive return. Ticker DQ had a -62.6% return. Another notable observation is that the stock trade volumes were much higher in 2018. With that being said, since the return was lower, we can infer that the these stock prices generally went down in 2018. The International Energy Agency posted an article in 2019 with the headline 'Renewable capacity growth worldwide stalled in 2018 after two decades of strong expansion'. It seems as if major policy initiatives surrounding clean energy stalled in 2018 and CO2 emissions were on the rise.

![2018_stock_performance_2](https://github.com/asliwinski23/Stock-Analysis/blob/main/2018_stock_performance_2.png)

### Script Execution Time
I had refactored my original script. Below find comparison screenshots. The 2017 analysis had an 211% improvement in runtime. The 2018 analysis had a 205% improvement in runtime.

2017 Original Script:


![module_2017_runntime](https://github.com/asliwinski23/Stock-Analysis/blob/main/module_2017_runntime.png)


2017 Refactored Script:


![VBA_Challenge_2017](https://github.com/asliwinski23/Stock-Analysis/blob/main/VBA_Challenge_2017.png)


2018 Original Script:


![module_2018_runntime](https://github.com/asliwinski23/Stock-Analysis/blob/main/module_2018_runntime.png)


2018 Refactored Script:


![VBA_Challenge_2018](https://github.com/asliwinski23/Stock-Analysis/blob/main/VBA_Challenge_2018.png)

## Summary
* What are the advantages or disadvantages of refactoring code?

The most glaring advantage of refactoring your code is better runtimes. When a code is more precise and can run over code faster. In this analysis the runtime was reduced from iterating over the number of rows by how many tickers we are analyzing to iterating over all of the stock data once. Refactoring can also help those who may need to run your code understand what is happening easier as it is more concise. Lastly, refactoring is a great excercise for the programmer to think about the implementation of their code when it is refactored vs not refactored.

Refactoring code can be a huge disadvantage if a team is working on a tight deadline. Programmers should not do refactoring at the last minute because you may not have the time to test the refactored code before release, which can introduce bugs. A way to mitigate the risk of introducing bugs into the code is to test. Depending on the company, the labor cost spend refactoring a large chunk of code can have adverse affects on financials, and for minimal return sometimes. In most cases, refactoring involves being more conscious of memory storage. However, to make this green stock analysis code run faster we saved each ticker's trade volume and price data in separate arrays, whereas in the original code these three elements were stored into three constant space variables. So in turn, we increased the memory usage in the refactored code.

* How do these pros and cons apply to refactoring the original VBA script?

It definitely took me a lot longer to refactor my code than to write the original. Plus, it was a lot more frustrating to figure out how to compact and test my analyses. Luckily, I did not wait until the very last minute to work on this analysis, so the disadvantages were not as apparent in this refactoring. The improvement in runtime was well worth the time and effort put in.
