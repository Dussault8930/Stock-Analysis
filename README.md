# Stock-Analysis
  The Purpose of this analysis was to determine the profitability of stocks based on their performance throughout the year 2017 and 2018 using scripts to compile the information into an easy-to-read table, as well as determine the usefulness of refactoring scripts.

---
### 2017 Analysis
<img width="457" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/71742174/95704942-1a9a4b80-0c18-11eb-9b55-5325d70f84ef.png">

  During this year, all but one of the stocks performed well enough to return some money on investment. Some even went on to return more than double the investment!

### 2018 Analysis
<img width="426" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/71742174/95704954-22f28680-0c18-11eb-92ff-65c628d8b8b0.png">

  This year was vastly different from the year prior, only 2 of the stocks that we looked actually brought money back to its investors.

  The script that was used to get the analysis on these two sets of data had been refactored at this point and ended up being far more efficient. In the original script there were a number of loops that were repetitive and cost more time for the code to run. A few of my "and" statements that I had within some of my "if" statements were also found to be unnecessary and were excluded as well. Before refactoring, the script took my computer about 2.5 seconds on average to run and as you can see from this images, that time was cut down to about a fifth of what it used to be.

---
## Summary
  Some of the advantages of refactoring are that is can cut down on the amount of time it takes for a computer to run the script and ultimately makes the script more efficient. This also allows anyone who might view the script itself to better understand it if the "fat" has been trimmed off and there aren't as many useless steps that are being taken. The flipside of this however, is that you lose the opportunity to expand on a particular part of the code in your script if you're trying to make it as efficient as possible. A personal issue I ran into is trying to refactor the script I was working with without breaking it, which happened a few times and I imagine happens a lot when trying to optimize a script.
  The original VBA script I was working with had a lot of moving parts with loops being run that essentially covered the same ground a few times over. Refactoring the code definitely made it run more efficiently and allowed for me to add in some notes on some of the lines to let anyone who viewed it know what I was thinking when writing the code.
