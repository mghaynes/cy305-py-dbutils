# Access Database Utilities
The purpose of this project is to create utilities to assist in grading the 
USMA CY305 database project. These utilities are designed to work with 
Microsoft Access files (.accdb).

The basic idea of the utility is to compare two database objects (tables or
 queries) against each other. The utility can also return a grade based on
 results of a comparison.
 
The main utility is **DAOdbUtils.py**. This module only works on Windows OS.
An older version named **dbUtils.py** is OS independent. However, it is not
being maintained and has less functionality.


## Required Setup
1. Install the appropriate Microsoft Access Database Engine. The code was
tested using Microsoft Access 2013 and the 
[Microsoft Access Database Engine 2010](https://www.microsoft.com/en-us/download/details.aspx?id=13255).

2. Install the following python modules:
  + win32com.client (pypiwin32)
  + levenshtein (python-Levenshtein)
  + numpy (numpy)

*Note: the levenshtein module requires installation of Visual C++ Build
 Tools. Code tested with [Visual C++ 2015 Build Tools](http://landinghub.visualstudio.com/visual-cpp-build-tools).
 
  
## Quick Start
### Loading a database file 
Get metadata from a database file by creating a DataBase object. The object
is instantiated with a path to the database file.

Example: 
```python
SolnDB = DataBase(SolnDBPath)
StudentDB = DataBase(StudentDBPath)
 ```
 
 The database object contains metadata on all the tables and queries in the
  project. For example, to list all the table names in the database:
 ``` python
   print(SolnDB.TableNames)
 ```
 The wiki documentation contains a complete listing of available variables 
 and functions.
 
 ### Comparing Tables
 With the database metadata loaded, you can compare any two tables with the 
 *AssessTables* function. The first parameter entered is the reference table.
 The second parameter will be compared against the first parameter.
 Example:
 ```python
 table_assessment = AssessTables(SolnDB.Tables['Platoon'], StudentDB.Tables['Platoon'])
 ```
 The *AssessTables* function returns an instance of class *TableScore*.
 *TableScore* contains comparison values for elements of a table. The
 elements compared include primary keys, relationships (i.e. foreign 
 keys), field names, and more. The wiki documentation contains a complete 
 listing of available variables and functions.
 
 ### Scoring Tables
 Scoring is based on weighting each of the fields in the *TableScore* 
 instance. This can be done using the *AssignTableWeights* function.
 Example:
 ```python
  table_weights = AssignTableWeights(NameScore=.05, FieldNameScore=.05, FieldTypeScore=.1, SamePriKeysScore=.4,
                     Correct_Num_Rltns=.025, Fld=.075, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.025,
                     Integrity=.075)
```
A weight can be assigned to each field of the *TableScore* class.
The weights should (but don't have to) add up to 1. 

Finally, use the *ScoreTable* function to return the score for the
table. *ScoreTable* essentially multiples the value for each field
by the assigned weight and returns an overall percentage grade (as a
 decimal number). Example:
 ````python
score = ScoreTable(table_assessment, table_weights)   
````
Note: You would probably then want to multiply the returned score 
ratio by the possible points for that table to get a point value.

### Comparing and Scoring Queries
Comparing and scoring queries works the same was as for tables.
However, you would use functions *AssessQuery*, 
*AssignQueryWeights*, and *ScoreQuery* which function exactly 
like their table counterparts.

## Contact
If you have questions or would like to help in maintaining this repo,
 contact me at either malcolm.haynes@usma.edu or mghaynes@gatech.edu. 