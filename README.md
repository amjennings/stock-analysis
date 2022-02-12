# stock-analysis
**Overview of the Project**
###### The purpose of this analysis was to create a file that would generate an analysis of stocks. Depending on the year input into a message box when opened, the analysis would display the total daily volume and the yearly return for twelve different stocks. To create this analysis in Excel, Visual Basic was used to create a loop that would run through all the data in the spreadsheet and pull just the information from the specified stocks. The analysis involved conditional statements that would change the color of the yearly return cell to green or red depending how the stock performed.
##**Results**
###### As shown in the images below, most stocks had greater performance in 2017 than in 2018. In 2017, all stocks except for TERP had an increase in yearly return. TERP had a decrease in yearly return for both years.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/98051208/153727219-e7b34548-2ee3-4fea-80ed-d3aaa71fdaf1.png)

###### The refactored script was nearly six times faster than the original script in generating the output.

<img width="590" alt="2017runtime_original_script" src="https://user-images.githubusercontent.com/98051208/153727242-1d5e6927-aeb7-4ab0-85d3-422db754e264.png">

###### In 2018, only ENPH and RUN had an increase in yearly return, whereas the remaining stocks had a decrease in yearly return. 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/98051208/153727257-5ef39599-d67b-405c-9994-64637c30ec0c.png)

###### The refactored script was more than six times faster than the original script in generating the output. 

<img width="590" alt="2018runtime_original_script" src="https://user-images.githubusercontent.com/98051208/153727270-3e36d888-f55e-47ad-abf4-c942576f6e39.png">

##**Summary**
###### The purpose and advantage of refactoring code is to make the code more efficient. The analysis is produced quicker from a code that requires less memory. Additionally, refactoring code may also be useful to learn better strategies for writing code. A disadvantage to refactoring code can occur if changing the code results in errors when working parts of the code are broken. As it applies to the VBA script, there is less code in the refactored script and the output was generated quicker. The disadvantage of refactoring the VBA script was it did require extensive troubleshooting to complete the refactor. However, I learned more strategies for troubleshooting, and it provided opportunities to practice explaining the code I wrote.
