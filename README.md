# Excel-Keyword-Lookups
User guide: [PDF](https://github.com/ITAutomator/Excel-Keyword-Lookups/blob/main/M365%20Teams%20Policy%20Update%20Readme.pdf)    
Download: [ZIP](https://github.com/ITAutomator/Excel-Keyword-Lookups/archive/res/heads/main.zip)    
Website: [WWW](https://www.itautomator.com/Excel-Keyword-Lookups/)   
(or click the green *Code* button (above) and click *Download Zip*)  


In Excel, find keywords based on a lookup table

![image](https://github.com/user-attachments/assets/bbc4690b-9e7f-49ad-a41a-b17b605d228a)


 

# What this does
If you have a column with description data and you need to pick out keywords, it’s easy if there’s just one keyword you are looking for.  This solution uses a lookup table of keywords.


# How it works

The final form of the function is this  
=IFERROR(TEXTJOIN(",",TRUE,FILTER(tblKeywords[Keyword],ISNUMBER(SEARCH(tblKeywords[Keyword],[@Description])))),"")  

The formula uses these functions
Search, returns the position of the first match from the Description column, and an error if not found. We use IsNumber(Search), since we don’t care about position. This is just True if it’s found, False if not.

The key here is the use of arrays. Notice that the SEARCH(tblKeywords[Keyword], [@Description]) refers to the entire column of keywords tblKeywords[Keyword] which is an array.  So the returned value will be an array of positions.

Then IsNumber takes the array of positions and converts it to an array of True/False (Boolean) values.

Filter, takes the array of keywords and the array of Booleans and lines them up, returning only the keywords with true next to them – a filter of the array.

If multiple keywords are found, Excel will spill them into neighboring cells (if they are empty cells).  Since we don’t want this, we use TextJoin to return everything as a single entry with the keywords separated by commas.

Lastly, if no keywords are found, Excel will show an empty array error. To suppress this, we use the IfError function.

