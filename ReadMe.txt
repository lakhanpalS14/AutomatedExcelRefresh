Problem Statement 
Automate the process of refreshing Excel Pivots connected to heavy databases and multidimensional models, and keep retrying until the successful refresh happens    
 
Scenario 
In OCP CPM, we get some Adhoc requests on regular intervals to load/create Excel reports on massive datasets. When we refresh the data present in the report, we usually encounter errors such as locking conflicts, out of memory issues, loss of connection, query time-out etc. There are no absolute solution to these problems but to just retry 2-3 times, and it will refresh. But it eats a lot of developers time in retrying and monitoring. 

Solution 
1.	We have implemented a configurable C# application, that will pick up the Excels to be refreshed from a Source location and Save at the destination on a successful refresh 
2.	On a failed refresh, the tool will re-attempt 3 times, you can configure this number according to your need 
 
How to run the tool 
1.	Update config.json file with “SourceLocation” and “DestinationLocation” for the .xlsx files to be refreshed. Solution uploaded here. 
2.	Run RefreshPivotsExcel.exe file. 
3.	Once the tool runs successfully, you will find the refreshed excels at the destination location. 
4.	Check for the failed refreshes from the message displayed on the console.  
