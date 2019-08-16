# Data Verification and analysis
I was assigned to extract data from our database in order to construct a montly report for the three man pollutants. I wrote a program that would analyze the data and generate some graphs into an excel file. This project is done entirely in Python with Pandas and OpenPyxl as the main libraries.

# Biggest Difficulties
* On the server we have Pandas V0.13, which means that a lot of the modern functions does not exist such as converting the datafram into a datetimeobject.
- I solved that problem by transforming the datafram column into a list and manualing converting the list into datetime then inserting it back into the dataframe.
* Openpyxl is not a frequently used package. Thus, finding some specific information was hard.

# What I've Learned
* I learned a lot about Pandas and how to avoid countless forloop to do data manipulations.

# What I could've done differently
Looking back at this project, I could've sperated the PM2.5 from O3 and NO2 because they required additional steps to process. Seperating the code into two modules can make things a lot cleaner to look at. I can also reference the colunm letter and number at the start of the for loop instead of doing "get_column_letter()" every time. 
