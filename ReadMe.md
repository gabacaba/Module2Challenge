# Refactor VBA code and measure stock market performance

## Overview of Project

### Purpose
Steve is a recent finance graduate who wants to help his parents make a wise investment. His parents would like to invest in green energy as they are passionate about the topic. Since there are many forms of green energy to invest in, Steve is preparing a stock market assessment to suggest the best option for his parents.

The general purpose of this analysis is to help Steve conduct this assessment and determine the best investment opportunity for his parents. I’ll be using the stock market information of 12 green energy companies from 2017 and 2018. I’ll also be using the program *Visual Basic for Applications*, also known as **VBA** to conduct the assessment.

The specific purpose of this task, however, is to teach me how to program better. Since I’ve done most of the analysis through the instructions on module 2, now I must refactor the code and make the analysis run faster. This was a challenge because this is the first time I used VBA. But I’m happy to say that I’m satisfied with my work and feel honored to be sharing this final product with you. 

## Results

The first time I performed the analysis for all stocks either in 2017 or 2018, Excel took more than 2 seconds to run the program. The challenge consisted of refactoring the code to make it run faster. Since I didn’t know what made a program run slow, I decided to investigate about it. 
What I understood from my google search is that a VBA code runs slow when it must interact many times with the excel sheet or when there are many loops within the code. This was crucial to understand because otherwise I wouldn’t understand the logic behind the instructions of the challenge.

The data gathered information of the 12 green energy companies for the days the market was opened in 2017 and 2018. Column G shows the stock market price and column H shows the total stock volume that was traded during the day. Therefore, the analysis consisted of two components. First, of comparing the initial price and the final price of the stock. Second, of calculating the total volume traded within the year.

### The logic behind the initial code 
The initial code used loops to obtain the starting price, the final price and the total volume of each stock.

First, it created an array for all the 12 tickers. Then, it opened a loop that would search the starting price, the final price and the total volume for each ticker. It is important to mention that every time it would find the three values for each ticker, it would write them on the Excel sheet. Therefore, the loop contained two loops. The first one would select the ticker and the second one would search for the values of the current ticker. 

### The logic behind the refactored code
The refactored code also used loops but in a more strategic way. 
The aim of the refactored code was to interact less with the excel sheet and therefore make the program run faster. To obtain this, it would be wise to obtain the values of all tickers in only one search versus conduct 1 search for each ticker. 

#### Use of Arrays
According to the Microsof defition of an array, an array is a *single variable with many compartments to store values*. A typical variable has only one storage compartment in which it can store only one value, but an array can store multiple values. Therefore, we created 3 output arrays to store the starting price, the ending price and the total volume of each ticker.

            *'1b) Create three output arrays*
                  *Dim tickerVolumes(12) As Long*
                   *Dim tickerStartingPrices(12) As Single*
                   *Dim tickerEndingPrices(12) As Single*
 
#### Setting the initial value to zero
I wouldn’t have thought of setting the values of the total volumes to zero if it wouldn’t have been for the instructions. It makes sense once I think about it. But it wouldn’t have occurred to me until I ran the program. So the purpose of this loop is just to make sure all the total volumes values start at zero. 
            *''2a) Create a for loop to initialize the tickerVolumes to zero.*
                 *For i = 0 To 11*
                 *tickerVolumes(i) = 0*
                 *Next i*

#### Use a Ticker INDEX
Acording to Microsoft, an index returns a value or the reference to a value from within a table or range. Since the program is going to store the values for each ticker in an array variable, it would be sufficient to loop through the sheet only once. However, in order to tell the program to go from one ticker to the next, this code uses a **Ticker INDEX**. 

The ticker index variable is going to help us in two ways. 
1.	It will tell the Ticker Array the current ticker we’re searching for 
2.	It will tell the Volume array, and the Starting Price array and the Final Price array in which position to store the values.

##### VBA refactored script
The loop reads as following. (The code is in italics.) 

Start of the loop and loop over all the rows in the spreadsheet. The variable RowCount has counted how many rows there are in the spreadsheet. 
    *For i = 2 To RowCount*
Then we create three conditionals to see what to do with the current values of the rows where is positioned. 
 Conditional A / This conditional increases the volume for current ticker.
     *If Cells(i, 1).Value = tickers(tickerIndex) Then*
       *tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value*
    *End If*

Conditional B/ This conditional checks if the current row is the first row with the selected tickerIndex.
    *If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then*
    *tickerStartingPrices(tickerIndex) = Cells(i, 6).Value*
    *End If*
          
Conditional C/ This conditional checks if the current row is the last row with the selected ticker.
   *If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then*
    *tickerEndingPrices(tickerIndex) = Cells(i, 6).Value*
   *End If*
    
In order to move into the next ticker, we use a conditional that says “If the next row’s ticker doesn’t match, increase the tickerIndex”.           
      *If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then*
       *tickerIndex = tickerIndex + 1*
       *End If*
I love the logic behind this refactored version of the initial code. Seems quite elegant and so efficient. Please see the images check the time in which this program runs. It's more than 4 times faster than the original VBA script. 


## Summary

### The advantages and disadvantages of refactoring code in general

I believe refactored code is less complex and somewhat easier to understand. However, it might be difficult to think of a simpler way of doing things once you’ve figure out a solution. I see it like an editing job. I used to be a copy writer for a bank, and I remember that writing a script versus editing the script were two completely different tasks. Once was creative and required ingenuity, the other one was thorough and required deep thinking. So perhaps one of the disadvantages is that it could be very time consuming.

### The advantages and disadvantages of the original and refactored VBA script

I believe that the original VBA script followed a natural logic. Meaning that it took a ticker at a time and searched for the values of the current ticker. Whereas the refactored VBA script took into consideration the time and instead created a beautiful method to store the values of each ticker in an array. I think the problem with the original VBA script is that since it contained only one variable for starting price, one for ending price and one for volumes, it needed to display the results on the table at the end of each loop. However, the refactored VBA script anticipated this limitation by created an array that could store multiple values. Therefore, it could separate the process into two tasks. The first task was to acquire the data and the second task was to display the data into a table. By doing this it wasn’t necessary to have a loop within a loop and interact with the excel sheet multiple times. This reduced time and make the program more efficient. 
