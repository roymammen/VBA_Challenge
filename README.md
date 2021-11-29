# VBA_Challenge
1. Overview of Project: Explain the purpose of this analysis.
  Overview and Purpose: 
	  To help Steve analyze a data set of stock transactions to analyze for Steve's parents who are planning to invest in an "alternate energy" stock. 
    Also, to identify if there is a possible diversification coming out of analysis.
	  
    To complete this process, help Steve, write VBA scripts to summarize overall performance of several stocks in 2017 and 2018 and to identify if "DQ" stock ticker symbol is an okay investment. 
	  Qualify this assumption from overall stock performance in 2017 and 2018 as well as identify any new trends which may be hidden to help Steve's parents to investigate further.
    The data set and VBA script summary can help Steve's parents to take a more informed decision on investing into the future.

  Analysis:
	  Using the data set from 2017 and 2018 a VBA script was written that would summarize. 
	    a. stock ticker volumes
	      Volume of stock would give an idea of how much a stock is in demand and the lesser the volume over years, could mean that the stock importance could be reducing.
	    b. Stock Price performance
	      Stock starting price beginning of year vs stock ending price ending of year
	      Track this over 2 years in 2017 and 2018
	      Price performance would give idea of the stock trend in terms of performance, given a possibility of a sustaining tried to identify hidden gem.


2. Results: 
	a. Using images and examples of your code, compare the stock performance between 2017 and 2018, 
	  Overall "DQ" Stock volume has increased from 2017 to 2018 by a factor of 3. 
    But the stock return in 2018 is -88.9% though this has improved from -336.5%. Hence this is not a good recommendation.
	  As a speculator one can buy DQ stocks in small quantities, do an averaging over a longer period which can eliminate big market moves. 
    But not a sure shot with the last two year performance.
	
    In addition, there were some hidden gems identified which were marked in "green" and trend in volume is also increasing.	
    Two of these stocks are. 
	  VLSR 	: Volume and returns are steady and consistent year over year. 
	  RUN 	: Volume is close to 2x doubled in 2018 and price performance year over year is 450%.

	b. Execution times of the original script and the refactored script.
    The execution of the script are more or less linear.
	  2017 alone sheet was processed in 1.92 seconds.
	  2018 sheets were processed in 2.69 sec.
	  
    The script has no abnormal issues and should be reasonable and scalable given more data should be added.

3. In a summary statement, address the following questions.
	a. What are the advantages or disadvantages of refactoring code?
	  Advantage: Advantage of refactoring a code is to have scalability when more data gets added.
	  Refactoring should be integrated with ease to help process data when more data set of 2019 and 2020 should be added.
	  Also refactoring should make the code bug free because it worked with a data set and should be able to apply to other new data sets when added.

	  Disadvantage:  One major disadvantage of refactoring is plan for proper testing to avoid any bugs.
	
  b. How do these pros and cons apply to refactoring the original VBA script?
	  Pros: VBA can be used to achieve completing repetitive tasks on a data set to get meaningful results and summary.
          VBA is easy to implement given Microsoft support.
	  Cons: VBA editing and debugging is a bit difficult, but doable. 
        The run time errors are not straightforward and there is an extra amount of time spent to get a working solution.
