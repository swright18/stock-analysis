


# Stock Analysis
Analysis of Stock Data used to determine whether the stock "DQ" is a wise choice to invest in and what stock options have had positive returns and which did not. 

## Purpose
The purpose of the project was to help Steve determine whether his parents are making a wise choice by putting their entire portfolio into the stock "DQ."

## Analysis
### Analyzing DQ Stock
In order to determine whether "DQ" is a stock that will be a good investment, we analyzed the data from stocks in 2017 and 2018. By looking at the Total Daily Volume and Return of the "DQ" stock. To accomplish this, VBA was utilized to calculate these totals (Figure 1) and to create a table displaying these totals. Buttons were then created in order to choose specific years and display specific data from that year and to clear that data (Figure 2a & 2b). Although the data for "DQ" looks promising for 2017, the most recent data shows that "DQ" had a -63% return for 2018. This data shows that DQ may not be the best investment option for Steve's parents. 
###### Figure 1
![DQ_Analysis](https://user-images.githubusercontent.com/79758494/112706903-5139ea80-8e75-11eb-8e92-525b876d4528.PNG)
###### Figure 2a
![DQ_Table](https://user-images.githubusercontent.com/79758494/112706923-7af31180-8e75-11eb-9a51-41e0afed5a02.PNG)
###### Figure 2b

![DQ_Table_2](https://user-images.githubusercontent.com/79758494/112706967-cefdf600-8e75-11eb-983f-aff16070b61b.PNG)

### Analyzing All Stocks
Once it was determined that "DQ" may not be the wisest investment choice for Steve's parents, we wanted to provide them with data to show what stocks would be good to invest in. To do this, we looked at the Total Daily Volume and Return for each stock we have data for and compiled that data into a single table for each year with buttons to help calculate, format, and clear the data (Figure 3a-3c)
###### Figure 3a
![All_stocks_1](https://user-images.githubusercontent.com/79758494/112707340-9ad80480-8e78-11eb-8059-15b9dc0964a6.PNG)
![All_stocks_2](https://user-images.githubusercontent.com/79758494/112707341-9ca1c800-8e78-11eb-98e8-2d1e8cc3bb7f.PNG)
###### Figure 3b
![All_Stocks_2017](https://user-images.githubusercontent.com/79758494/112707235-e342f280-8e77-11eb-9765-df7e1fe8d861.PNG)
###### Figure 3c
![All_Stocks_2018](https://user-images.githubusercontent.com/79758494/112707243-edfd8780-8e77-11eb-8654-d5c6fe441dbf.PNG)
### Refactoring Stock Analysis Code
After we accomplished this goal, Steve asked how this VBA code would work on analyzing much larger Stock Data Sets. We refactored the code in Figure 3a in order to analyze the data quicker and more effectively. The code in Figure 4 shows the refactored code. 
###### Figure 4
![Refactor_1](https://user-images.githubusercontent.com/79758494/112707495-ee971d80-8e79-11eb-9957-ef3e9139fea3.PNG)
![Refactor_2](https://user-images.githubusercontent.com/79758494/112707531-27cf8d80-8e7a-11eb-9ba5-35e3a75850db.PNG)
### Refactoring Results
In order to determine whether the refactored code ran more efficiently than the other, we put a Time Start and Time End in order to calculate how long the code ran before it completed. We also created a message box to pop up with that time to show how long it took to run. Figure 5a and 5b show the time it took before refactoring. 
###### Figure 5a
![2017 - Before](https://user-images.githubusercontent.com/79758494/112707609-a62c2f80-8e7a-11eb-877b-8e30c3034c1f.PNG)
###### Figure 5b
![2018 - Before](https://user-images.githubusercontent.com/79758494/112707611-a88e8980-8e7a-11eb-9a6e-e64bfd42d8c8.PNG)

We did the same after refactoring the data (Figure 5c and 5d) 
###### Figure 5c
![VBA_Challenge_2017](https://user-images.githubusercontent.com/79758494/112707613-acbaa700-8e7a-11eb-86e8-dbf7b880448a.png)
###### Figure 5d
![VBA_Challenge_2018](https://user-images.githubusercontent.com/79758494/112707617-afb59780-8e7a-11eb-894c-fa352148c1fb.png)

As the figures about show, after refactoring the data, the Code ran approximately 5 to 5.5 times faster. Although with the amount of data we have, the time seems insignificant, with larger data sets, this could matter significantly. 

## Advantages and Disadvantages of Refactoring
- A disadvantage of refactoring code is how long it may take to do it. On this code specifically, it did not seem too tedious, but if one is working with hundreds and hundreds of lines of code, it could be using resources that may not be available to write code that already works. 
- The advantages of refactoring code is having simpler, quicker code that is easier to read. Although the code may take a while to refactor, at the end, the code is going to be clearer and possibly more effective at doing its job. 
## Advantages and Disadvantages of Refactoring Stock Analysis VBA Code
- The disadvantage of refactoring the Stock Analysis code was how long it took. 
- The advantages of refactoring the Stock Analysis code was how much quicker the code ran following the refactoring. Also, by refactoring, it is going to be able to run much larger data sets quicker and is going to be able to be used for years to come due to the "yearValue" added. 