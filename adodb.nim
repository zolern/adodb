#====================================================================
#
#               adodb - Microsoft ADO DB accessor
#                (c) Copyright 2019 Encho "Zolern" Topalov
#
#====================================================================

##[
This module provides methods for access ADO DB compatible databases. 
SQL string literal interpolation and timestamp literal processing
also included. Some useful routines, like NZ or DSum are provided also.

Nim implementation was strongly inspired from nodeJS npm module 
[node-adodb](https://github.com/nuintun/node-adodb) and uses almost 
the same idea to access ADO DB with query and exec methods.

(adodb uses [winim](https://github.com/khchen/winim) for low-level
access to Microsoft ADO)

Usage:
   
.. code-block:: Nim
   import adodb

   # Connect to Microsoft Access 2003 DB
   let connect = r"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\test.mdb"
   let adoDb = newADODB(connect)

   # Execute SQL statement in db
   adoDb.exec("DELETE * FROM Users WHERE ((UserID)>100)")

   let user = "Test User"
   let id = 1
   
   adoDb.exec(sql"INSERT INTO Users ( user_id, user_name ) VALUES ({id}, '{user} 1')")
   adoDb.exec(sql"UPDATE Users SET user_bday=#31.12.1999# WHERE ((user_id)={id})")
   
   # Retrieve data from db
   let rst = adoDb.query("SELECT * FROM Users")

   # Access data
   for rowIdx, row in rst:
      for idx, name, fld in row.fields:
         echo "Row ", rowIdx, ", Fld[", idx," | ", name, "] = ", fld

To compile:

.. code-block:: Nim

   nim c source.nim
      add -d:winansi or -d:useWinAnsi for Ansi version (Unicode by default)
      add -d:notrace disable COM objects trace. See com.nim for details.
      add -d:useWinXP for Windows XP compatibility.

Database connection strings:
For Access 2000-2003 (\*.mdb): 
   .. code-block:: Nim
      connect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<path-to-mdb-file>;"

For Access > 2007 (\*.accdb): 
   .. code-block:: Nim
      connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<path-to-accdb-file>;Persist Security Info=False;"
]##

{.deadCodeElim: on.}

import tables
import winim/com as wincom
import os
import unicode

import adodb/private/sqlformat
import times

export wincom, sqlformat, times

type
   ADODB* = ref object
      connection: string

   ADOField = ref object
      data: variant

   ADODataRow = seq[ADOField]

   ADORecordset* = ref object
      fields: OrderedTableRef[string, int]
      data*: seq[ADODataRow]

   ADORow* = ref object
      parent: ADORecordset
      data: ADODataRow

# ADOField method & properties
proc init(self: ADOField, v: variant) =
   self.data = copy(v)

proc final(self: ADOField) =
   if not self.data.isNil:
      self.data.del
   self.data = nil

proc value(self: ADOField): variant =  copy(self.data)

proc newADOField(v: variant): ADOField =
   new(result, final)
   result.init(v)

# ADORecordset properties and methods
proc init(self: ADORecordset, rst: com) =
   self.fields = newOrderedTable[string, int]()

   if rst.BOF() != -1 or rst.EOF() != -1:
      let fields = rst.Fields()

      for fld in fields:
         let fldName: string = fld.Name()
         self.fields[fldName.toLower()] = self.fields.len

      rst.MoveFirst()
      
      while rst.EOF() != -1:
         self.data.add(@[])
         for fld in fields:
            self.data[self.data.len - 1].add(newADOField(fld.Value()))
         rst.MoveNext()

proc close*(self: ADORecordset) =
   ## Release memory allocated for recordset data.
   ## It is a good practice to call recordset's 'close' 
   ## when its data is not needed anymore
   while self.data.len > 0:
      var row = self.data.pop
      while row.len > 0:
         var fld = row.pop
         fld = nil

proc final(self: ADORecordset) =
   self.close

proc newADORecordset(rst: com): ADORecordset =
   new(result, final)
   result.init(rst)
         
proc len*(self: ADORecordset): int {.inline.} = self.data.len
   ## returns row count of ADORecordset

template rowCount*(self: ADORecordset): int = self.len
   ## alias of len

# ADORow methods & properties
proc init(self: ADORow, rst: ADORecordset, data: ADODataRow) =
   self.parent = rst
   self.data = data

proc newADORow(rst: ADORecordset, data: ADODataRow): ADORow =
   new(result)
   result.init(rst, data)

proc `[]`*(self: ADORecordset, index: Natural): ADORow =
   ## return index-th (from 0 to len-1) row of recordset 
   result = newADORow(self, self.data[index])

iterator items*(self: ADORecordset): ADORow {.inline.} =
   ## iterates through recordset rows
   for row in self.data:
      yield newADORow(self, row)

iterator pairs*(self: ADORecordset): (int, ADORow) {.inline.} =
   ## iterates through rows, returns pair (index, ADORow)
   for rowIdx, row in self.data:
      yield (rowIdx, newADORow(self, row))

# ADORow properties and method
proc `[]`*(row: ADORow, index: Natural): variant = row.data[index].value
   ## returns value by index of column/field

proc `[]`*(row: ADORow, field: string): variant =
   ## returns value by name of column/field
   let index = row.parent.fields[field.toLower()]
   return row.data[index].value

proc len*(row: ADORow): int {.inline.} = row.data.len
   ## returns count of fields

template fieldCount*(row: ADORow): int = row.len
   ## alias of len

iterator items*(row: ADORow): variant {.inline.} =
   ## iterates all columns of row
   for fld in row.data:
      yield fld.value

iterator pairs*(row: ADORow): (int, variant) {.inline.} =
   ## iterates all columns, return pair (index, value)
   for fldIdx, fld in row.data:
      yield (fldIdx, fld.value)

iterator fields*(row: ADORow): (int, string, variant) {.inline.} =
   ## iterates all columns, return (index, field name, value)
   let fields = row.parent.fields
   for fldName, fldIdx in fields.pairs:
      yield (fldIdx, fldName, row.data[fldIdx].value)

proc release(o: var com) =
   if o.isNil: return

   if o.State() != 0:
      o.Close()
   
   o = nil

# ADODB methods and properties
proc query*(adoDb: ADODB; sql: string): ADORecordset =
   ## retrieve recordset by SQL statement

   var rst: com = nil

   try:
      # Create recorsed
      rst = CreateObject("ADODB.Recordset")

      # Open recordset
      rst.Open(sql, adoDb.connection, 0, 1)

      # Query data
      result = newADORecordset(rst)
   finally:
      rst.release()

proc exec*(adoDb: ADODB; sql: string) =
   ## execute SQL statement

   var conn: com = nil
   try:
      # Create connection
      conn = CreateObject("ADODB.Connection")

      # Open
      discard conn.Open(adoDb.connection)

      # Execute
      discard conn.Execute(sql)
   finally:
      conn.release()

proc init(self: ADODB, connection: string) =
   self.connection = connection

proc newADODB*(connection: string): ADODB {.inline.} =
   ## Constructor
   new(result)
   result.init(connection)

# Variant useful routines
template isNull*(v): bool =
   ## checks variant is null
   when v.type is variant:
      (v.isNil or v.rawType == 1)
   elif compiles(v.isNil):
      v.isNil
   else:
      false

template nz*(v, valueIfNull): untyped =
   ## returns v if v is not null, otherwise returns valueIfNull
   (if isNull(v): valueIfNull else: v)

# Several data aggregation methods

proc dQuery(adoDb: ADODB; functor, field, domain, where: string): variant =
   # base procedure, that is used for all aggregate functions
   result = nil
   var statement = sql"SELECT {functor}({field}) AS {functor}Of{field} FROM {domain}"
   if where != "":
      statement.add sql" WHERE ({where})"

   let rst = adoDb.query(statement)
   
   if rst.rowCount != 0 and rst[0].fieldCount != 0:
      result = rst[0][0]
   
   rst.close

proc dMax*(adoDb: ADODB; field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return Max value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "Max", field, domain , criteria)

proc dMin*(adoDb: ADODB; field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return Min value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "Min", field, domain , criteria)
      
proc dFirst*(adoDb: ADODB; field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return first value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "First", field, domain , criteria)
   
proc dLast*(adoDb: ADODB; field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return last value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "Last", field, domain , criteria)
      
proc dLookup*(adoDb: ADODB; field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "", field, domain , criteria)

proc dCount*(adoDb: ADODB; field, domain: string; criteria: string = ""): int =
   ## return count of rows by field in domain, 0 if no data in domain
   ## (optional criteria can be applied)
   nz(dQuery(adoDb, "Count", field, domain , criteria), 0)

proc dSum*(adoDb: ADODB, field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return sum of all values of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "Sum", field, domain, criteria)

proc dAvg*(adoDb: ADODB, field, domain: string; criteria: string = ""): variant {.inline.} =
   ## return average value of all values of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   dQuery(adoDb, "Avg", field, domain, criteria)
   