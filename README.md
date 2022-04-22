# The VBA of Wall Street

VBA scripting was used to analyze generated stock market data. The [Test Data](https://docs.google.com/spreadsheets/d/1R11XYttakv2qdKjh9vnxryQBTnogepQp/edit?usp=sharing&ouid=101500589655888677304&rtpof=true&sd=true) file was used while developing the script and the script was used to run on the [Stock Data](https://drive.google.com/file/d/1ij29MGBPUCORnyZUx0POns2UmhkjGOsd/view?usp=sharing) file to generate the final report.  




The created script loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

Functionality was added to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". Also, appropriate adjustments were made to the VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.

