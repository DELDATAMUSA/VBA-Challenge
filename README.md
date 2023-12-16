# VBA-Challenge
# VBA-Challenge

Code explanation

1) The stock Analysis subroutine is defined.This subroutine should perform stock analysis for each worksheet in the workbook.
2)The subroutine initializes variables and sets the header labels for each worksheet in cell L1 ("ticker"),cell J1 ("Yearly Change"),cell K1 ("Percent Change") and cell L1("Total Stock Volume")
3)The LastRow variable is set to the last row of data in Column A(Ticker symbol) for the current worksheet
4)The SummaryRow variable is set to 2,indicating the starting row for writing the summary information (excluding the header row).
5)A loop is initiated from row 2 to the last row of data column A.
6)The code checks if the ticker symbol in the current row is different from the next row. If it is different ,it means we have reached  the end of a stock's data for the year
7)The opening price,closing price and ticker symbol for the current stock are assigned to variables.
8)The yearly change is calculated by subtracting the opening price from the closing price.
9)The percentage change is calculated by dividing the yearly change by 100.If the opening price and multiplying by 100.If the opening price is zero,the percentage change is set to zero to avoid division by zero error.
10)The ticker symbol,yearly change,percentage change,and total stock volume are writtern to the summary rows in columns I,J,K AND L respectively.
11)The percentage change value is formated as a percentage with 2 decimal places.
12)Conditional formatting is applied to highlight positive changes(green) & negative changes (red) in the yearly change column.
13)The summary row index is incremented, and the total volume is reset to zero for the next stock.
14) The total volume for each stock is calculated by summing up the volume values in column G.
15) After the loop finishes for the current worksheet,the code proceeds to find the stock with the greatest percentage increase.
16) The max percent change is assined the maximum value from the percentage change column.
17)The ticker symbol corresponding to maximum percentage change is determined to find the row number.The row number is obtained by adding 1 to the row  number.
18)The maximum percentage change and its corresponding ticker symbol are writtern in P2 AND Q2 cells respectively.
19)Similarly the code identifies the stock with the greates total volume and the total volume are written to cells P4 & Q4 respectively.
21)The loop continues for the next worksheet, and the process is repeated until all worksheets have been analysed.
