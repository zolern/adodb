# Module adodb

adodb enables simple access to Microsoft ADO compatible databases. It is inspired from 
[node-adodb](https://github.com/nuintun/node-adodb). adodb provides "out-of-the-box" 
SQL literal interpolation ({} fields, similar to strformat module) and timestamp literal
processing. Some well known aggregation functions like DSum, DMin, DMax, etc. are also provided

adodb internally uses awesome [winim](https://github.com/khchen/winim) for low-level access to 
Microsoft ADO and COM processing.

## Code Example

```nimrod
import adodb

# Connect to Microsoft Access 2003 DB
let connect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\test.mdb"
let ado = newADODB(connect)

# Execute SQL statement in db
let id = 100
ado.exec(sql"DELETE * FROM Users WHERE ((UserID)>{id})")

# Retrieve data from db
let rst = ado.query("SELECT * FROM Users")

# Access data
for rowIdx, row in rst:
   for idx, name, value in row.fields:
      echo "Row ", rowIdx, ", Fld[", idx," | ", name, "] = ", value
```

If the code above is saved to sample.nim you can compile it, as follows:

    nim c sample.nim

## Construction, destruction, properties and methods of adodb

Constructor:

```nimrod
let ado = newADODB(connection)
```

Database connection strings:
- For Access 2000-2003 (\*.mdb): 
	
```nimrod
connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<path-to-mdb-file>;"
```

- For Access > 2007 (\*.accdb): 
	
```nimrod
connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<path-to-accdb-file>;Persist Security Info=False;"
```

Methods:

    exec - execute SQL statement, no data is returned
    query - execute SQL statement and returns ADORecordset object

## Methods of ADORecordset

Methods:

	len/rowCount - return recordset's row count
	[index] - return row by index (from 0 to len - 1)

Iterators:
	
	items - iterates through recordset rows (return ADORow)
	pairs - iterates through recordset rows, return (index, ADORow)

## Methods of ADORow

Methods:

	len/fieldsCount - return row's field count
	[index] - return value of field by index
	[name] - return value of field by name

Iterators:

	items - iterate through fields, return values
	pairs - iterate through fields, return (index, value)
	fields - iterate through fields, return (index, name, value)

## Deal with NULL variants

Two useful templates are added to manage fields with no value (variant.VT_NULL)

    isnull - check if variant is NULL
    nz - replaces value NULL with appropriate value

## Data aggregation functions

AdoDB provides several useful data aggregation methods:

```
	dMin (field, domain [, criteria]) is equal to SELECT Min(field) FROM domain [WHERE criteria])
	dMax (field, domain [, criteria])     ...     SELECT Max(field) FROM domain [WHERE criteria]
	dSum (field, domain [, criteria])     ...     SELECT Sum(field) FROM domain [WHERE criteria]
	dCount (field, domain [, criteria])   ...     SELECT Count(field) FROM domain [WHERE criteria]
	dFirst (field, domain [, criteria])   ...     SELECT First(field) FROM domain [WHERE criteria]
	dLast (field, domain [, criteria])    ...     SELECT Last(field) FROM domain [WHERE criteria]
	dLookup (field, domain [, criteria])  ...     SELECT field FROM domain [WHERE criteria]
```	

## Install
With git on windows:

    nimble install https://github.com/zolern/adodb

Without git:

    1. Download and unzip this module (by click "Clone or download" button).
    2. Start a console, change current dir to the folder which include "adodb.nimble" file.
       (for example: C:\nim\adodb>)
    3. Run "nimble install"

For Windows XP compatibility, add:

    -d:useWinXP

## Documents
    
   * [adodb](https://zolern.github.io/adodb/adodb.html)
   * [sqlformat](https://zolern.github.io/adodb/sqlformat.html)
    
## License
Read license.txt for more details.

Copyright (c) 2019 Encho Topalov, Zolern. All rights reserved.
