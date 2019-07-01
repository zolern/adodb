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
import threadpool
import os
import unicode
import segfaults
import times

import winim/com
import adodb/private/sqlformat

export com, sqlformat, times

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

   ADOResult = ref object
      rst: ADORecordset
      err: ref Exception

# ADOField method & properties
proc init(self: ADOField, v: variant) =
   self.data = copy(v)

proc done(self: ADOField) =
   if not self.data.isNil:
      self.data.del
   self.data = nil

proc final(self: ADOField) =
   self.done

proc value(self: ADOField): variant =  copy(self.data)

proc newADOField(v: variant): ADOField =
   new(result, final)
   result.init(v)

# ADORecordset properties and methods
proc init(self: ADORecordset, rst: com) =
   self.fields = newOrderedTable[string, int]()

   if rst.BOF() != -1 or rst.EOF() != -1:
      let fields = rst.Fields()
      let fieldCount = fields.Count()

      rst.MoveFirst()
      while rst.EOF() != -1:
         self.data.add(@[])
         for i in 0 ..< fieldCount:
            let fld = fields.Item(i)
            
            if self.fields.len < fieldCount:
               let fldName: string = fld.Name()
               self.fields[fldName.toLower()] = self.fields.len
      
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
         fld.done
         fld = nil

proc final(self: ADORecordset) =
   self.close

proc newADORecordset(rst: com): ADORecordset =
   new(result, final)
   result.init(rst)
         
proc len*(self: ADORecordset): int = self.data.len
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

iterator items*(self: ADORecordset): ADORow =
   ## iterates through recordset rows
   for row in self.data:
      yield newADORow(self, row)

iterator pairs*(self: ADORecordset): (int, ADORow) =
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

proc len*(row: ADORow): int = row.data.len
   ## returns count of fields

template fieldCount*(row: ADORow): int = row.len
   ## alias of len

iterator items*(row: ADORow): variant =
   ## iterates all columns of row
   for fld in row.data:
      yield fld.value

iterator pairs*(row: ADORow): (int, variant) =
   ## iterates all columns, return pair (index, value)
   for fldIdx, fld in row.data:
      yield (fldIdx, fld.value)

iterator fields*(row: ADORow): (int, string, variant) =
   ## iterates all columns, return (index, field name, value)
   let fields = row.parent.fields
   for fldName, fldIdx in fields.pairs:
      yield (fldIdx, fldName, row.data[fldIdx].value)

proc release(o: var com) =
   if o.isNil: return

   if o.State() != 0:
      o.Close()

   o.del
   o = nil

proc thQuery(connection, sql: string): ADOResult =
   {.gcsafe.}:
      ## retrieve recordset by SQL statement
      new(result)

      var rst: com

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
         CoUninitialize()

proc query*(adoDb: ADODB; sql: string): ADORecordset = 
   let resFlow = spawn thQuery(adoDb.connection, sql)
   let res = ^resFlow
   
   if not res.err.isNil:
      raise res.err

   return res.rst


proc thExec(connection, sql: string): ADOResult =
   ## execute SQL statement
   {.gcsafe.}:

      ## retrieve recordset by SQL statement
      new(result)

      var conn: com = nil
      
      
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
   let resFlow = spawn thExec(adoDb.connection, sql)
   let res = ^resFlow
   
   if not res.err.isNil:
      raise res.err

proc init(self: ADODB, connection: string) =
   self.connection = connection

proc newADODB*(connection: string): ADODB =
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

proc dQuery[T](adoDb: ADODB; functor, field, domain, where: string): T =
   var statement = sql"SELECT {functor}({field}) AS {functor}Of{field} FROM {domain}"
   if where != "":
      statement.add sql" WHERE ({where})"

   let resFlow = spawn thQuery(adoDb.connection, statement)
   let res = ^resFlow
   
   if not res.err.isNil:
      raise res.err

   result = fromVariant[T](res.rst[0][0])
   res.rst.close

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
   