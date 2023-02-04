Flow:

Text to column 
Swap columns 
pre max version-- sort alphabateically A to z both version and ID
Max version formula A2=A1 
Data time format (will also have section)
Table Filters--basic EDF basic filters + does not contain fesseinheim + only "1" for latest

At this point, the neccesary table is ready
Now we need to split the data into three different tables
and in the end combine that into a single sheet ready to be copy pasted into the lynx copy


lynx copy format

Current Outages
<br>
Reactor Name            Capacity (MW)       Outage Start        Expected Start      Reason
<br>
Entries                 MROUND(v,100)       mmm. dd, yyyy       mmm. dd, yyyy       Statutory/Planned/Unplanned  
<br>
Total Outages           Sum(CurrentOut)     
<br>
Planned Outages in future 
<br>
Entries                 MROUND(v,100)       mmm. dd, yyyy       mmm. dd, yyyy       Statutory/Planned/Unplanned
till upto 6 months from today
<br>
Recent Restarts
<br>
Entries Latest 13                 MROUND(v,100)       mmm. dd, yyyy       mmm. dd, yyyy       Statutory/Planned/Unplanned   