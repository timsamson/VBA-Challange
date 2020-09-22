<h1># VBA-Challange</h1>

VBA script created to read through exsisitng stock data to extract the diiference beween the opening and closing prices for the stock for the year. 
Additional steps to find the stock on each worksheet with the highest and lowest change in rate and greatest trading volume. The script goes thru each workbook and looks for the assocaited data comparing stock tickers until it finds one taht os out of order. It's imparative  that the sheets are sorted by stock ticker before running the macro otherwise the script will nto provide the correct results. Note: the Stock market data provided was sorted in such a way.

<b>Notes on script.</b>
<ul><li>On Error needed to be added to the script to combat and error 6 messsage on line 56 of the code, this was due to divisible by 0 error when teh script ran near the end of each worksheet.</li>
  
<li>Ideally the Total_Volume integer would like to  have been set to a double but was changed to variant as there is a know issue with Mac on excel and this was the only workaround i could find that provided consisitent results.</li></UL> 
  
<b>BOTH of these changes may not be required if running on a  windows based machine. </b>
  
<h3>Scripts (located in script folder):</h3>
  
  <ul><li> MyMacro_VBA_Wallstreet.vbs <b> (please use for grading)</b></li></ul>
  Working Files:
  <ul><li>Sample_data_script.bas used while working on debugging and creating the initial script.</li>
  <li>VBA_Wallstreet_Loop.bas This is the working version of the script.</li></ul>
  
<h3>Screenshots (Located in Screenshot Folder):</h3>
 
2014_VBA_WS.png

2015_VBA_WS.png

2016_VBA_WS.png

<h3>Additional Notes</h3>
Sample_Data_Script contains the nessecary calls to "ActiveWorksheet" to facilitate the loop to run thru various worksheets, but does not conatain the nessecary "For" statement to finalize that process. Additonally it contains some errors related to formatting and column types that was lated resolved in the final script "VBA_Wallstreet_Loop.bas"


