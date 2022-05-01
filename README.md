# Stock Analysis Overview

The original goal of this analysis was to assist the client in assessing the dataset of the stock market for the last few years to understand trends in stock profitability. Once that was completed, the new goal was to refactor the code I had created to make my code run more efficiently with the goal that it could handle larger amounts of stock data beyond the dozen I used for this project. No additional functionality was added to the code, I simply streamlined it to better fit the client’s needs and to make it easier to read and understand for any future users. 

# Results

## Stock Performance Comparison

(insert images here)
In the above images, thanks to the red and green fillers, it is easy to see which stocks performed well and which did not. This quick comparison also allows us to see that 2017 was a much better year for the majority of these 12 stocks than 2018 and having the data for both years easily accessible will make it simple for the client to draw conclusions about the best path forward for their investments. Even I, with no experience with stocks, can see that EMPH and RUN were the only two stocks to turn a profit in both 2017 and 2018, which logically makes them the best candidates for invested based on this data alone (other factors may influence that decision, of course). 

## Execution Times of Original Vs. Refactored Scripts
(insert images here)
These are the run times for the original code. 
(insert images here)
These are the run times for the refactored code. It is clear that the refactored code runs quite a bit faster than the original and since it produces the same desired results, this would be the code best suited for the clients needs and could also be more widely applicable for larger data sets. 

# Summary

## What are the Advantages or Disadvantages of Refactoring Code? 

### Advantage:
One advantage is on clear display in the previous section. The purpose of refactored code is for it to run much faster. This also leads to the next logical assumption, since the refactored code is faster it can process larger data sets. This is beneficial for our clients in general, as the more data we can apply the more informed of a decision they can make. 

Another advantage is the fact that refactored code is more streamlined than the original, which means it is easier to read and easier to adjust for any specific needs. With less components that require editing, it is easier to continuing refactoring this code to fit different needs with less risk of bugs that will render it ineffective. 

### Disadvantage:
A potential disadvantage is that the original code had more failsafe functions built into it. While those functions slowed the code down, they also made sure that the code properly process through all the data. The refactored code functioned perfectly for the current data set, but if there was a large amount of data, there may be errors that presented themselves because of the lack of appropriate failsafe functions. 

## Applying these Advantages and Disadvantages to the Original and Refactored VBA Script
When determining which code to provide to the client, it is important to compare the advantages and disadvantages to ensure that I give them the best tool for the job at hand. In this situation, the advantages clearly outweigh the disadvantages. The refactored code is faster and more efficient in analyzing the dataset I was given to work with. It will also be easier for the client to adjust the refactored code if they need to because it is simplified. The potential disadvantage present only applies if a much larger data set was used and again the previous point applies here as well, even if the refactored code struggles with the larger data set, it will be easier to make the necessary adjustments to ensure it can still handle the task. 
