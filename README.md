VBA SQL writer for Excel to Query Access Database
------------
Add this module to your Excel VBA project to quickly set up SQL statements for an Access Database.  Especially when using forms with multiple Command Buttons, each running a different query, this module allows you to set up the query just using Arrays.

Features
---------
Run SQL - SELECT, INSERT, UPDATE, or DELETE statements by passing in the file location, table name, and a series of arrays.  The module functions will write and run the query.  In the case of SELECT, an Array will be returned.

Examples
--------
```VBA
Private Sub runSELECT()

  Dim result() As Variant
  accessFile = "C:\file location\file.accdb"
  dbTable = "someTable"              'Name of table in Access
  lastName = ComboBox1.Value
  colArray = Array("firstname")      'Array of all the columns that you want returned  
  whereArray = Array("lastname")     'Array of all the columns names you want to use to select result  
  compareArray = Array(lastName)     'Array of inputs that you want to compare to columns to select result. Add wildcards for partial search: Array("%" & lastName & "%") 
  oppArray = Array("=")              'Array of operators for comparing, such as "LIKE", "<", or "<"
                                     'You can use as many criteria for selection as you like as long as the where, compare, and opperator arrays have the same number of / and properly ordered items.
                                     
  result = sqlSELECT(accessFile, colArray, dbTable, whereArray, compareArray, oppArray)

'receiving the result as an Array requires a little different handling than just using the data set while you have an open connection, but it can easily be tailored to your needs.

'Other Function Examples: (Insert, Update, and Delete return a Boolean, TRUE if the query was successful)
Dim result() As Boolean
result = sqlInsert(accessFile, dbTable, colArray, valArray)
result = sqlDelete(accessFile, dbTable, whereArray, compareArray, oppArray)
result = sqlUpdate(accessFile, dbTable, colArray, updateArray, whereArray, compareArray, oppArray)

```


