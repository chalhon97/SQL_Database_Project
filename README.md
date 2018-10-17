## SQL_Database_Project

This project was for my database management systems course where I created a database for a business, and helped them migrate it to SQL sever via an ODBC driver. This allowed for the benefits of a network based DBMS while still allowing one to have Microsoft access functionality and control (if one desires). 

In addition, the database was then imported to Microsoft Excel by setting up a connection string and using a mix of VBA and SQL to program the macros to do specific functions. (Please see specific macro files for further details). This allows for easier databse manipulation and analysis. 

## Further Explanations:

The ODBC file was used to connect the Microsoft Access database to Microsoft SQL Server. 

Opening the access file you should see that each table has been recognized as a database object (DBO) and that is the table which is linked to SQL server. This allows to the database to be stored on the cloud but still allows for control and manipulation through Microsoft Access.

# IMPORTANT:
The macros require Microsoft ActiveX data objects 6.1 to run (currently unavailable for mac). Please make sure this is enabled when you attempt to run the macros otherwise you will get a compile error. 

To enable this open the macro window and go to tools -> references and check off the ActiveX data objects 6.1 box if not     already checked. 
  
Even if you are unable to run the macros feel free to look at the macro code files seperately.
