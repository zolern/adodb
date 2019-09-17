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
         echo "Row ", rowIdx, ", Fld[", idx," | ", name, "] = ", $fld

To compile:

.. code-block:: Nim

   nim c source.nim
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

import os, threadpool, times
import unicode, tables
import json except `$`

import winim/[winimx, com]
winimx currentSourcePath()

import adodb/private/[sqlformat, parsevariant, jsontime]

export sqlformat, times
export json except `$`

type
   ADODB* = ref object
      connection: string

   ADOField = JsonNode

   ADODataRow = seq[ADOField]

   ADORecordset* = ref object
      fieldNames: OrderedTableRef[string, int]
      rows*: seq[ADODataRow]

   ADORow* = ref object
      parent: ADORecordset
      fields: ADODataRow

   ADOResult = ref object
      rst: ADORecordset
      err: ref Exception

# ADORecordset properties and methods
proc init(self: ADORecordset, rst: com) =
   self.fieldNames = newOrderedTable[string, int]()

   if rst.BOF() != -1 or rst.EOF() != -1:
      rst.MoveFirst()
      let fieldCount = rst.Fields().Count()
      
      while rst.EOF() != -1:
         let fields = rst.Fields()
   
         self.rows.add(@[])
         
         for i in 0 ..< fieldCount:
            let 
               fld = fields.Item(i)
               fldName: string = fld.Name()
               fldLowerName = fldName.toLower()

            if self.fieldNames.len < fieldCount:
               self.fieldNames[fldLowerName] = self.fieldNames.len
      
            self.rows[self.rows.len - 1].add(variantToJson(fld.Value()))
         rst.MoveNext()

proc newADORecordset(rst: com): ADORecordset =
   new(result)
   result.init(rst)
         
proc len*(self: ADORecordset): int = self.rows.len
   ## returns row count of ADORecordset

template rowCount*(self: ADORecordset): int = self.len
   ## alias of len

# ADORow methods & properties
proc init(self: ADORow, rst: ADORecordset, row: ADODataRow) =
   self.parent = rst
   self.fields = row

proc newADORow(rst: ADORecordset, data: ADODataRow): ADORow =
   new(result)
   result.init(rst, data)

proc `[]`*(self: ADORecordset, index: Natural): ADORow =
   ## return index-th (from 0 to len-1) row of recordset 
   result = newADORow(self, self.rows[index])

iterator items*(self: ADORecordset): ADORow =
   ## iterates through recordset rows
   for row in self.rows:
      yield newADORow(self, row)

iterator pairs*(self: ADORecordset): (int, ADORow) =
   ## iterates through rows, returns pair (index, ADORow)
   for rowIdx, row in self.rows:
      yield (rowIdx, newADORow(self, row))

# ADORow properties and method
proc `[]`*(row: ADORow, index: Natural): ADOField = row.fields[index]
   ## returns value by index of column/field

proc `[]`*(row: ADORow, field: string): ADOField =
   ## returns value by name of column/field
   let index = row.parent.fieldNames[field.toLower()]
   return row.fields[index]

proc len*(row: ADORow): int = row.fields.len
   ## returns count of fields

template fieldCount*(row: ADORow): int = row.len
   ## alias of len

iterator items*(row: ADORow): ADOField =
   ## iterates all columns of row
   for fld in row.fields:
      yield fld

iterator pairs*(row: ADORow): (int, ADOField) =
   ## iterates all columns, return pair (index, value)
   for fldIdx, fld in row.fields:
      yield (fldIdx, fld)

iterator fields*(row: ADORow): (int, string, ADOField) =
   ## iterates all columns, return (index, field name, value)
   let fields = row.parent.fieldNames
   for fldName, fldIdx in fields.pairs:
      yield (fldIdx, fldName, row.fields[fldIdx])

proc value*[T](fld: ADOField): T =
   when T is DateTime:
      result = jsontime.getDateTime(fld)
   elif T is string:
      result = fld.getStr()
   elif T is BiggestInt:
      result = fld.getBiggestInt()
   elif T is SomeInteger:
      result = fld.getInt()
   elif T is SomeFloat:
      result = fld.getFloat()
   elif T is bool:
      result = fld.getBool()

proc `$`*(fld: ADOField): string = 
   when T is DateTime:
      result = `$`(fld.value[DateTime])
   else:
      result = json.`$`(fld)

converter adoFieldToDateTime*(fld: ADOField): DateTime = value[DateTime](fld)
converter adoFieldToString*(fld: ADOField): string = value[string](fld)
converter adoFieldToInt*(fld: ADOField): int = value[int](fld)
converter adoFieldToBiggestInt*(fld: ADOField): BiggestInt = value[BiggestInt](fld)
converter adoFieldToFloat*(fld: ADOField): float = value[float](fld)
converter adoFieldToBool*(fld: ADOField): bool = value[bool](fld)

proc release(o: var com) =
   if o.isNil: return

   if o.State() != 0:
      o.Close()

   o.del
   o = nil

proc thQuery(connection, sql: string): ADOResult =
   {.gcsafe.}:
      ## retrieve recordset by SQL statement
      var rst: com
      
      new(result)

      try:
         CoInitialize(nil)

         rst = CreateObject("ADODB.Recordset")
     
         # Query
         discard rst.Open(sql, connection, 0, 1)
     
         # Get data
         result.rst = newADORecordset(rst)
         
      except:
         result.err = getCurrentException()

      finally:
         rst.release()
         COM_FullRelease()
         CoUninitialize()

proc query*(adoDb: ADODB; sql: string): ADORecordset = 
   let 
      resFV = spawn thQuery(adoDb.connection, sql)
      res = ^resFV
   
   if not res.err.isNil:
      raise res.err

   return res.rst

proc thExec(connection, sql: string): ADOResult =
   ## execute SQL statement
   {.gcsafe.}:

      ## retrieve recordset by SQL statement
      var conn: com = nil

      new(result)
      
      try:
         CoInitialize(nil)
         # Create connection
         conn = CreateObject("ADODB.Connection")
         
         # Set CursorLocation
         conn.CursorLocation = 3

         discard conn.Open(connection)

         # Execute
         discard conn.Execute(sql)

      except:
         result.err = getCurrentException()

      finally:
         conn.release()
         COM_FullRelease()
         CoUninitialize()

proc exec*(adoDb: ADODB; sql: string) =
   let 
      resFV = spawn thExec(adoDb.connection, sql)
      res = ^resFV
   
   if not res.err.isNil:
      raise res.err

proc init(self: ADODB, connection: string) =
   self.connection = connection

proc newADODB*(connection: string): ADODB =
   ## Constructor
   new(result)
   result.init(connection)

proc dQuery[T](adoDb: ADODB; functor, field, domain, where: string): T =
   var statement = sql"SELECT {functor}({field}) AS {functor}Of{field} FROM {domain}"
   if where != "":
      statement.add sql" WHERE ({where})"

   let 
      resFV = spawn thQuery(adoDb.connection, statement)
      res = ^resFV

   if not res.err.isNil:
      raise res.err

   return res.rst[0][0].value[:T]
      
proc dMax*[T](adoDb: ADODB; field, domain: string; criteria: string = ""): T =
   ## return Max value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[T](adoDb, "Max", field, domain , criteria)

proc dMin*[T](adoDb: ADODB; field, domain: string; criteria: string = ""): T =
   ## return Min value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[T](adoDb, "Min", field, domain , criteria)
      
proc dFirst*[T](adoDb: ADODB; field, domain: string; criteria: string = ""): T =
   ## return first value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[T](adoDb, "First", field, domain , criteria)
   
proc dLast*[T](adoDb: ADODB; field, domain: string; criteria: string = ""): T =
   ## return last value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[T](adoDb, "Last", field, domain , criteria)
      
proc dLookup*[T](adoDb: ADODB; field, domain: string; criteria: string = ""): T =
   ## return value of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[T](adoDb, "", field, domain , criteria)

proc dCount*(adoDb: ADODB; field, domain: string; criteria: string = ""): int =
   ## return count of rows by field in domain, 0 if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[int](adoDb, "Count", field, domain , criteria)

proc dSum*(adoDb: ADODB, field, domain: string; criteria: string = ""): float =
   ## return sum of all values of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[float](adoDb, "Sum", field, domain, criteria)

proc dAvg*(adoDb: ADODB, field, domain: string; criteria: string = ""): float =
   ## return average value of all values of field in domain, null if no data in domain
   ## (optional criteria can be applied)
   result = dQuery[float](adoDb, "Avg", field, domain, criteria)
   