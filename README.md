# VBA-challenge

This VBA file is written to do the following:

First, add column headers: "Ticker," "Yearly Change," "Opening Price," "Closing Price," "Percent Change," and "Total Stock Volume" to the spreadsheet.

Then, by identifying the last row of the dataset, and by identifying interruptions in the repetition of ticker symbols, the script adds each unique ticker symbol to the column Ticker.

Next, the script identifies the opening price located next to each new ticker symbol, and the closing price located next to each final occurrence of each last ticker symbol, and places these values in the appropriate columns "Opening Price" and "Closing Price."

The script then calculates the net change in Ticker symbol price over the course of the year, and the corresponding percent, and adds these values to the appropriate columns.

Finally, the script color codes percent change as green for positive change, and red for negative change. 