# product ID matcher 
Key terms:
Shortage = items that were never received in full from the vendor.
Overage = items that exceeded initial purchase order from vendor. 
P.O = Purchase Order
Record = row
Field = column
ROM = Receipt Overage Management Database 

Variance report = report sent to the vendor indicating the details of every product that was ordered. 
This Macro was designed to improve the accuracy, functionality and speed of all variance reports created for Zulily's Vendor Operations Team. For the macro to work properly, purchase order data must be imported from Tableau into a macro-enabled Excel spreadsheet. Once the data is in the Excel Spreadsheet the appropriate macro can be activated. In the developer ribbon, you will notice 2 modules. Each module extends the same goal: transfer shortages into overages. It's that simple; however, each module is used in different circumstances. The first module is used when manual P.O's are created by a vendor operations specialist to group existing overages that were miss-interpreted by the database as shortages. These overages are indicated with a blank P.O Push date (which makes sense because the items were never actually ordered on the original P.O date). So for each record that contains an empty push date its value is transferred from the "P.O shortage" field to the "P.O Overage" field. Once `sub transfer()` is complete, sub main() calls sub format(), which, as the name suggest, formats the existing data into Zulily-specific formatting scheme (vendor name, P.O #, Product ID, SKU, Vendor SKU, Desc., Size, Cost, Shorted %, P.O qty, rcvd qty, P.O under, P.O over). For instance, some fields are irrelevant, and subsequently deleted by the macro. 

note: The shortcut key for module 1 is crtl + shift + "v". 

Once the data is formatted then the second module can be run. The goal of this module is the exact same as the first module, except it compares product ID's from the ROM data to the current variance report product list to find which records need to have adjustments made. By adjustments, I am referring to transfers, such as shortages into overages, or overages that need to be inserted from the ROM into the variance report. The program logic is as follows. First, the data from the ROM is imported into a worksheet adjacent to the variance report data (so "Sheets2"). It contains two fields of interest, "product ID" and "overage qty". The program oscillates back and forth between the two sheets to find a product match. It looks only in the fields designated as "product ID". If a match is found, the overage quantity from Sheet2 is transferred to the variance report (Sheet1) in the "overage qty" field, and the process repeats itself until all the product ID's have been searched. 
note: the shortcut key for module 2 is crtl + shift + "n".
