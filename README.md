<h1># VBA-Challange</h1>

VBA script created to read through exsisitng stock data to extract the diiference beween the opening and closing prices for the stock for the year. 
Additional steps to find the stock on each worksheet with the highest and lowest change in rate and greatest trading volume.

<b>Notes on script.
  On Error needed to be added to the script to combat and error 6 messsage on line 56 of the code, this was due to divisible by 0 error when teh script ran near teh end of each worksheet.</b>
  
  Ideally the Total_Volume integer would like to  have been set to a double but was changed to variant as there is a know issue with Mac on excel and this was the only workaround i could find that provided consisitent results. 
  
  BOTH of these changes may not be required if running on a  windows based machine. 
  
<h3>Scripts (located in script folder):</h3>
  
  Sample_data_script.bas used while working on debugging and creating the initial script.
  VBA_Wallstreet_Loop.bas This is the working version of the script. (please use for grading)
  
<h3>Screenshots (Located in Screenshot Folder):</h3>
 
2014_VBA_WS.png

2015_VBA_WS.png

2016_VBA_WS.png

