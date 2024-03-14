# VBA-challenge
VBA Challenge

This VBA code analyzes stock data from multiple years. The following steps were followed to develop the code:

1. Declared the variables to be used in the code.
2. Defined the range of the data, starting from the first column and row A1, up to the last row identified in the data.
3. Defined the location of the first ticker in the data, as well as the first unique ticker variable.
4. The loop iterates through each row, identifying if the ticker exists or is new.
5. If the ticker is new, it prints the name and calculates the yearly change between the last close amount and the first open amount.
6. Similarly, it calculates the percent change between the last close amount and the first open amount.
7. The loop continues, and the values are stored until the code finds a new unique ticker. The stored values are for the last row (i) and the first row (j) associated with the unique ticker.
8. The total volume is calculated during the loop, and at each row, it sums the cumulative volume and the volume from the next row, only if it corresponds to the same ticker since it is still in the same loop.
9. After the loop finalizes, an additional summary table is displayed for the greatest increase, greatest decrease, and greatest volume among the unique tickers.
10. The last row of the unique tickers table is defined.
11. A loop that goes through the unique tickers' data will find the greatest values by comparing with the previous cell, for both increase, decrease, and volume.
12. These results are printed on an additional table.
13. Final formatting is provided for number format, autofit of cells, and headers of each column.