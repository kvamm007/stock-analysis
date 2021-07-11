# Module 2 Challenge: VBA and Refactoring Code

## Purpose of Project

The purpose of this project was to create the same output that we had created working through the module to learn Visual Basic, but with code that runs more efficiently. The more efficient code also gave us additional practice with creating arrays and utilizing output arrays. 


## Results

### Time/Efficiency Difference
The refactored code ran significantly faster; averaging the two times for 2017 and 2018, the refactored code ran in 0.2246 seconds. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85597801/125208257-12c71080-e257-11eb-901a-7db32a6ae5bb.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85597801/125208264-16f32e00-e257-11eb-92fb-b1458260ac4b.png)

The original code created while learning the module, averaging the two times for 2017 & 2018, ran in 1.3633 seconds. 

![2017 Original Time](https://user-images.githubusercontent.com/85597801/125208279-21152c80-e257-11eb-96e6-25a54df2bf5f.png)
![2018 Original Time](https://user-images.githubusercontent.com/85597801/125208284-25414a00-e257-11eb-9eb9-11bd1a2edd48.png)


While this is only a difference of 1.1387 seconds, which would not be very noticeable in most cases, this can add up if the processes are being run repeatedly; it may also make a greater time difference on a computer with less processing power, or a computer busy processing other tasks using resources. 

If looked at as a factor, rather than flat time, then the refactored code is approximately 6 times more efficient, which is a big difference. 



### Stock Performance in 2017 and 2018

To measure stock performance, we used a macro to run through the list of all stocks, and for each ticker, total the trading volumes for the year. We also found the initial closing price and the last closing price for each ticker, and the percentage difference for the year is the percent gained or lost. To make it easier to read, we added formatting to auto-fit column width, display the volumes in a number format with commas, and highlight values green (positive) or red (negative). The largest part of the code is summarized & displayed below under challenge specific refactoring. 

#### Overall Market Comparison
The ending results of stock performance between the two years for the 12 selected stocks is drastically different. The highlighting of negative and positive values makes the contrast extremely stark- there was only 1 stock which lost value over the course of 2017, while only 2 stocks gained value in 2018. DQ in particular is a stark contrast, gaining 199.2% in 2017 (highest percentage gain) while losing 62.6% in 2018 (highest percentage of loss). 

![2017 Results- Closing](https://user-images.githubusercontent.com/85597801/125208322-39854700-e257-11eb-9c18-ddfa13970a1d.png)
![2018 Results- Closing](https://user-images.githubusercontent.com/85597801/125208305-325e3900-e257-11eb-8474-92bd76b44452.png)

#### Recommended Stock
Based on analysis of just these two years, and these 12 selected stock tickers, ENPH is likely the best stock to bet on; it had a 129.5% increase in 2017 (3rd highest) and a 81.9% increase in 2018 (2nd highest, one of only two to be positive). Its trading volumes also increased from around 220 million in 2017 to 610 million in 2018. If trading volume can be seen has a measure of interest in the stocks, then its interest increased dramatically between 2017 and 2018, and of the stocks we analyzed, it had the highest trading volume in 2018. 


### Another Analysis: First opening value to last closing value
#### Coding Adjustment
A slight adjustment to the way to view the stocks would be to go from the first trading of the stock *opening* value, rather than closing value. Using the first closing value to the last closing value essentially removes the first day of the year from the analysis. Only a tiny coding tweak is needed to adjust this:

![Change Code to Opening Price](https://user-images.githubusercontent.com/85597801/125208337-47d36300-e257-11eb-80fa-002fb85c305c.png)


In the code to set the “tickerstartingprices” value for each ticker index, we merely adjust Cells(i,6) to be Cells(i,3). Rather than take the closing value, from the 6th column, we take the opening value from the 3rd column. The results for this are as follows:

![2017 Results- Opening](https://user-images.githubusercontent.com/85597801/125208340-4bff8080-e257-11eb-805d-041e52749b29.png)
![2018 Results- Opening](https://user-images.githubusercontent.com/85597801/125208494-69811a00-e258-11eb-9fc1-0acc0d65e8cd.png)



#### Resulting Difference

The differences are not large enough to make a difference in the recommendation in this case; however, you can see that the ENPH stock, when using the opening price (truly from the start of 2018) has an even greater increase in 2018 of 97.9% (versus 81.9% with closing price) and is now the highest percentage increase for the year. None of the stocks were a big enough difference to change from positive to negative (or vise versa) and I would still recommend the same stock based on this data, however, it is interesting to see the difference. I would also argue that initial opening value to final closing value is the true analysis of price change for the year.


## Advantages and Disadvantages of Refactoring Code

### Advantages

#### General Advantages
Refactoring code, as we saw with the time differences above, can make a significant difference in the performance time. While this was not a huge difference in terms of actual time in this case, if we were to want to do this for the entire stock market, for example, a 6x factor on time would likely make a very large difference in time. 

In general, refactoring allows us to revisit code we had written previously when we were less experienced, or review code that someone else has written, and make it more efficient. If this was a very long and time consuming code to write, say thousands of lines, it would definitely be worth reviewing and refactoring small parts of the code as needed. 


#### Challenge Specific Advantages

With this challenge, by refactoring, we were able to make the code run more efficiently. While there is no significant difference in the length of the code (we were not able to make it significantly shorter, as we had not been repeating code in the first place) it did make a difference in performance. In the original code:

```
    '4) loop through tickers
        For i = 0 To 11
    
            ticker = tickers(i)
            totalvolume = 0
            
            '5) Loop through rows in the data
                Worksheets(yearValue).Activate
                
                For j = 2 To RowCount
                
                '5a) Get total volume for current ticker
                    If Cells(j, 1).Value = ticker Then
    
                        totalvolume = totalvolume + Cells(j, 8).Value
        
                    End If
                
                '5b) Get starting price for current ticker
                     If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
 
                        startingPrice = Cells(j, 6).Value
                        
                     End If
                
                '5c) Get ending price for current ticker
                    If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

                        endingPrice = Cells(j, 6).Value
 
                    End If
                
                Next j
                
            '6) Output data for current ticker
                    Worksheets("All Stocks Analysis").Activate
                    Cells(4 + i, 1).Value = ticker
                    Cells(4 + i, 2).Value = totalvolume
                    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


        Next i

```
What this code does is to go to the first row of data, see if the ticker is the specified value (“AY” in this example), if the ticker is “AY” then increase the total volume. Then, if the ticker is “AY” and the previous row is NOT “AY”, then set the current row’s closing price as the “startingprice” for the ticker. Then, if the ticker is “AY” and the next row is NOT “AY”, then set the current row’s closing price as the “endingprice” for the ticker. This will loop through every row in the spreadsheet, checking for “AY”, once it reaches the bottom, it will output “AY” ticker’s data into the results spreadsheet. It will then start again at the top with the next ticker, “CSIQ”. This means it will fully loop through all 3,000+ rows of data 12 times. 

With the refactored code:
```
For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
            If Cells(i, 1).Value = tickers(tickerindex) Then
            
                tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value
                
            End If
        
            
        '3b) Check if the current row is the first row with the selected tickerIndex.

              If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
 
                tickerstartingprices(tickerindex) = Cells(i, 6).Value
                        
              End If
            
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
            
              If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
              
                tickerendingprices(tickerindex) = Cells(i, 6).Value
                        
              End If
            
            

            '3d Increase the tickerIndex.
             If Cells(i + 1, 1).Value <> tickers(tickerindex) Then
             
                tickerindex = tickerindex + 1
                
            End If

```

The code essentially follows the same process of going to a row, checking to see if the ticker matches (“AY” to begin with), increasing the volume if it does, seeing if the previous row was NOT “AY” and then setting the starting value if yes, and checking to see if the next row was NOT “AY” and setting the ending value if no. The difference is that we use an index for the ticker here, rather than a nested loop, and once we see that the next row ticker is NOT equal to the current ticker (we have reached the end of the “AY” rows), we increase the ticker index by 1, to move on to the next ticker. Doing this, we only have to loop through the 3,000+ rows of data a single time. Once all rows have been looped through, then all data is output at once into the results sheet. This saves 11 times of looping through over 3,000 rows of data. We were able to do this with additional skills we learned with setting up arrays, that we were not familiar with when we first wrote the code. 


### Disadvantages

#### General Disadvantages

I am not sure personally how much I would use refactoring of code, even my own code. In my current job, I frequently write reports with a system that we had migrated to at the start of 2019. I will now often open reports I had originally written when we were brand new to the system and see that they were horribly written (I was new to the software and had no idea what I was doing) and rather than try to edit them, I find it is easier to just start over from scratch. However, most of those reports are things that take me less than 5 minutes to make with my current skill level- so, perhaps with a larger project, it would be worth going through and editing what already exists. This system also does not easily allow for comments to be inserted within the code as with VBA, which makes it much harder to figure out what you were doing or thinking at a specific point in the project. 

In general, I also think that as a new coder it is very difficult to figure out refactoring code- not writing your own from scratch and figuring it out as you go, but jumping into something fully or partially written by someone else. This may be easy when you are very experienced, but any time you are with a newer program or language, it’s going to be very difficult. 

#### Challenge Specific Disadvantages

In the case of the original learning module, where we went through step by step, I was able to easily follow the steps, and check at each step to see if everything seemed to be working properly. With refactoring the code, it was not nearly as easy to do in small steps and check the work, as I was having a harder time following what I was doing. I ended up writing all the code, and then having an error, which I very much struggled to figure out. I had failed to connect the last step of increasing the ticker to the tickerindex, so the tickerindex very quickly increased above the 11 variables set in the array, and it was difficult to figure out why I was getting an error. While it was a good learning experience and it did make the code faster, it was also frustrating and time consuming. In the end, it saved slightly over 1 second of time to run- it would take a long time to make up the 90 minutes I spent refactoring the code with 1 second of additional efficiency (it would need to be run 300,000+ times, which doesn’t seem likely). If it takes more time to refactor than you gain in efficiency, then it is not a good use of your time. 
