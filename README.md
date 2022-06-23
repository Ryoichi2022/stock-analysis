Challenge 2

# stock-analysis
## 1. Overview of Project
A code was developed using Visual Basic for Applications (VBA) to analyze transaction volume and rate of return for 12 stocks during the period of 2018 and 2017. This project is to refactor the code that has been created for the stock analysis in order to gain more efficiency in running the code. 

## 2. Results
### i. Stock data provided for analysis
The stock data was provided in Excel format and included opening price, highest price, lowest price, closing price, adjusted closing price, and transaction volume for each ticker (stock symbol) and transaction date of the year. In each year, the data consisted of 3012 lines, which means 251 business days of data for 12 stocks.

### ii. Overview of original code
In the original VBA code, all 12 tickers were defined as array, and the code was written using the For loop so that the program would run through the data for every one of the 12 tickers separately. A nested loop was adopted as shown below.
![Original_Code](https://github.com/Ryoichi2022/stock-analysis/blob/main/Original%20Code.png)

The process time resulted in 1.351563 seconds and 1.296875 seconds for the data of 2018 and 2017, respectively.

### iii. Overview of refactored code
In the original code, it appeared that the program repeated the same process over the same data 12 times. Therefore, there was room for refactoring to eliminate the repetition to gain more efficiency. Specifically, the code was edited as follows.
![Refactored_Code](https://github.com/Ryoichi2022/stock-analysis/blob/main/Refactored%20Code.png)

In the refactored code, the program runs only once over the stock transaction data. Every time the ticker changes in the date while the program is scanning it from the top to the bottom, the program switches the ticker to the next based on tickerIndex.

### iv. Result of refactoring
Improvement was seen in the process time, which was 0.21875 seconds and 0.21875 seconds for 2018 and 2017, respectively under the refactored code. For both years, the process time decreased by over 80% in the refactored code. This would be a significant difference when the program runs over much larger amount of data.

![Result_2018](https://github.com/Ryoichi2022/stock-analysis/blob/main/VBA_Challenge_2018.PNG)
![Result_2017](https://github.com/Ryoichi2022/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

## 3. Summary
### i. Advantage of refactoring code
The most important aspect of refactoring is efficiency. As the project demonstrated, a successful refactoring will cut the program’s process time by a certain amount. Also, by refactoring in this project, the code was able to get rid of nested structure, which is often complicate and not easy to understand.
As a learner, it creates more knowledges and experiences on programing or offers a chance to think in a different way.

### ii. Disadvantage of refactoring code
On the other hand, it will take time to refactor a code. I will work sometimes, but the other not. Some people would say, “it is fine to keep using the original code, as long as it works.” Refactoring may not create as much efficiency as the amount of time the programmer has needed on it. Comparison between the amount of time that is saved and the amount of time that is spent would be even more critical from a business the business viewpoint.

### iii. How refactoring worked in the project
Due to the volume of the data to be processed, efficiency that was gained by refactoring is quantitatively small, that is approximately one second. The disadvantage exceeds advantage in this refactoring project. However, I certainly learned a new way to build a code through the project.
