# The VBA of Wall Street

VBA scripting was used to analyze generated stock market data.



The following data files were used:

* [Test Data](https://docs.google.com/spreadsheets/d/1R11XYttakv2qdKjh9vnxryQBTnogepQp/edit?usp=sharing&ouid=101500589655888677304&rtpof=true&sd=true) - Used while developing scripts.

* [Stock Data](https://drive.google.com/file/d/1ij29MGBPUCORnyZUx0POns2UmhkjGOsd/view?usp=sharing) - Ran scripts on this data to generate the final report.



A script that loops through all the stocks for one year and outputs the following information was created:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

Functionality was added to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

Appropriate adjustments were made to the VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.



