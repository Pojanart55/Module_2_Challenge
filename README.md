# Module_2_Challenge
VBA Homework
Below please find my thought logic behind the code for this challenge
1. Data Structure and Assumptions:
  * The stock data is organized in multiple worksheets, with each worksheet represeting a quarter.
  * Each row within a worksheet represents a single day's data for a specific ticker symbol.
  * The goal is to analyze each worksheet (quarter) independently and calculate key metrics for each ticker within that quarter.
2. Algorithm:  here's a breakdown of the code highlighting the key functionalities
  * Looping Through Worksheet: Iterates through each worksheet in the workbook
  * Process Each Worksheet:
      * Within each worksheet:
          * Determine the last row with data in the Ticker column.
          * Loop through each row in the worksheet.
              * Check if the current row's ticker is differnt from the next row's ticker. If different...
              * Calculate the quarterly change, the percentage change
              * Calcuate and store the total stock volume for the current ticker.
              * Update the greatest percentage increase and decrease values if applicable.
              * Update the greatest total volume and its corresponding ticker if applicable.
              * Reset the total stock volume for the next ticker.
              * If the same then .. add the current row's volume to the total stock volume.
  * Display results
      * Display the calculated values in the designated cells within each worksheet.
      * Display the greatest percentage increase, decrease and total volume for each worksheet.
      * Autofitting columns for better readability.
  * Key Considerations
      * Data integrity: ensure that all values are numeric and correctly formatted for performing the script.
      * Error handling: I spent many hrs in debugging. I utilized Xpert Learning Assistant and Gemini throuhgout the debugging process. I have to admit that it's quite stressful when the output results didn't display as expected. I have learnt a lot though. 

 Attached please find the VBA file, Screenshots of the results for each quarter.  Please note that the screenshot only display the results from the summary table up to row 30. Therefore, I submit link to the Excel Macro file as part of my submission as well. 
https://drive.google.com/file/d/1869DlGNXHpmw379nqmb3Rbch_2PBO9DK/view?usp=sharing

