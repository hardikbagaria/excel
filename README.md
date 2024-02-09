# excel
here are some Visual Basics Marcos code I wrote with 7 modules inside it to peform these tasks:
1: auto create a pdf in ths specific code the VBA uses the ThisWorkbook.Sheets("Sheet1").Range(" ").Value to get party name and
   bill number and then creates a pdf named as <Bill No & Party Name & ".pdf">
2: Automatically send the file to a predefined number on whatsapp; tho this number can also be dynamic and can get value form 
   the excel file using ThisWorkbook.Sheets("Sheet1").Range(" ").Value and then sent to the python code using Shell command along 
   side the file name,
   The Automatic Sending is Carried out by a Python Program which uses selenium to send the file to the defined number 
   (Selenium can also be used to send the same to emails both pre defined or dynamic according to user needs)
3: It can also Automatically print the Bills And change the Right Header to "Original" and "Duplicate" (you can also add Triplicate and more)
4: There is Another Module For writing numbers in word called SepllNumbers
5: There is one more Code Which is used for automatically getting the next Bill No. it is run when the Workbook is opened and it checks 
   the Directory where all the bills are saved then uses RegEx to get the Number from the file name (As mentioned earlier bills are saved
   starting with the bill number) it compared it ang gives the highest value, now we add 1 to it and then use 
   Sheets("Sheet1").Range("G2").Value = maxInvoiceNumber to set it automatically when the Workbook is loaded
6: The most intresting of all is A Module to Add the Specific bill to Ledger. It is done in 2 Ways a common file for all of the nammed ledger.xlsx
   and And the second method is it adds it seperatly to dedicated party name file, it uses an if else statment to check weather a file of that name
   exists or not if not it creats one and then adds the data retriven using ThisWorkbook.Sheets("Sheet1").Range(" ").Value to the seprate excel workbook
   (its still a bit not organized cause it appends the name of the party in the second case also which is not required) At then end it gived a MsgBox that the data is added Successfully
7: i have used =IFERROR(VLOOKUP(C7,Sheet2!A:I,2,0),"") IFERROR is used if there is no data for the specific party name (i.e. Address Line 3 is not there for some Parties)
   it shows a blank rather than an error VLOOKUP uses the party name to check for the corrospondind data in Sheet2 it is done for GST no Address Cnt Person Email Ph. No etc..
8: along side all these it hax more functionalities like =TODAY() for automatic date automatic calculation, automatic GST calculation, automatic round off etc.
