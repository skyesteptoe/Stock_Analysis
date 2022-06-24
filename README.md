# Stock_Analysis
# Refactoring Code to Speed-up Stock Analysis

## Overview of Project
Steve is a recent college graduate with a degree in Finance. His parents are eager to take advantage of Steve's education and have asked Steve to review their current stock market investments. Our initial task was to write code for Steve so that he could compare the performance of his parent's current investment to that of similar companies in renewable energy. Steve quickly realized the power of this code; he could analyze his entire dataset with the click of a button. This was exciting but Steve was eager to look at more expansive datasets including stocks from other companies. Steve asked us to review our code and to see if updates would be needed in order to apply it to larger datasets. We decided to refactor the code with two goals: 1) optimizing run-time for the current yearly analyses and 2) updating our conditional statements so that they could be leveraged and applied more readily to other stock datasets.  

### Purpose
The purpose of this exercise is to rewrite our previous stock analysis code with the goal of improving or speeding up run-time. If the run-times are reduced with the modified code, the intention is to apply this refactored code to larger stock performance datasets. Further, the refactored code will be written to reduce the number of modifications required should other stock performance datasets be analyzed. 

## Analysis and Challenges

The refactoring analysis was completed by adding three additional output arrays to the previous yearly stock analysis code. These arrays were then used in the nested loops and were indexed to a pre-set "tickerIndex" of 0. In the original code, an index was not used and the variables were referenced directly in the conditional statements rather than via the arrays. Optimally, the only change that that will be necessary with the inclusion of other stock performance datasets will be updating the ticker list and the total number of arrays.  The intention is that this code can be applied more expansively as it is indexed to a specific reference and uses arrays that can be applied to larger datasets.

### Challenges and Difficulties Encountered
This exercise was challenging. The concept of indexing the three output arrays to the tickerIndex was not intuitive to me and the writing of the IfThen statements was complex. I only overcame this challenge by attending the TA sessions ahead of class and reviewing other shared code from classmates. I found myself making the looping statements overly complicated and the IfThen not complex enough. 

I finally thought I'd completed my code and when I tried to run it, I had to debug. I couldn't figure it out and then, Excel crashed with all my tweaks. Finally, I stepped away from the code and came back in the morning to find that my error was simple. I'd included spaces in my reference to my Excel worksheet that didn't have any spaces. I was relieved it was something simple in the end. I only realized this though by slowly running my code, line by line. 

## Results and Conclusions
In conclusion, this refactored code runs much more efficiently and should be used as Steve grows his financial advising business. It is often said that time is money and this refactored analysis will add up to time saved for Steve. He'll have the ability to readily advise his clients with rapid analysis of stock performance datasets.

Comparatively, the original code was between 4 and 7 times slower compared to the refactored code. At the human level, this is nearly imperceptible but nevertheless, these fractions of seconds add up over time. In the real world, these small segments of time add up to a lot of additional trades on the stock market where much of trading is done via automation.

While this code was more complex to write than the original, it is code that can be applied to larger datasets and by referencing arrays that can readily be added to (e.g., adding more tickers/stocks) without going to the trouble of re-writing other components of the code. The time spent upfront will make this code faster and easier to deploy as more stocks are included in Steve's analysis. 

A potential limitation to this refactored code and to the original code is the requirement to input each ticker by acronymn and assign a number. This seems a manual process that could likely be done in a more automated fashion using another program. Another limitation is writing this in VBA which may not be as intuitive to younger programmers.

With this dataset, you could look at a variety of descriptive statistics comparing stocks and their value per share. You could run similar analysis but instead look at other variables instead of just daily close. For example, you could use adjusted daily close or open price. For this particular dataset, it might be interesting to look at external trends and see if some stocks are more responsive to things related to sustainable energy - perhaps oil crises, gas prices, or chain supply shortages for solar panel components cause fluctuations for these stocks.
