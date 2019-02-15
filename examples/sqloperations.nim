import adodb

let 
   connect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb"
   ado = newADODB(connect)

   prefix = "Test User"
   id = 1

   tbl = "Users"

ado.exec(sql"DELETE * FROM {tbl}")

ado.exec(sql"INSERT INTO {tbl} ( user_id, user_name ) VALUES ({id}, '{prefix} {id}')")
ado.exec(sql"INSERT INTO {tbl} ( user_id, user_name ) VALUES ({id + 1}, '{prefix} {id + 1}')")
ado.exec(sql"INSERT INTO {tbl} ( user_id, user_name ) VALUES ({id + 2}, '{prefix} {id + 2}')")
ado.exec(sql"INSERT INTO {tbl} ( user_id, user_name ) VALUES ({id + 3}, '{prefix} {id + 3}')")

let updateSQL = sql"UPDATE {tbl} SET"
ado.exec(sql"{updateSQL} user_bday=#31.12.1999# WHERE ((user_id)={id})") # dd.mm.yyyy
ado.exec(sql"{updateSQL} user_bday=#{1985, 8, 22} WHERE ((user_id)={id + 1})") # mm/dd/yyyy
ado.exec(sql"{updateSQL} user_bday=#1987-11-15# WHERE ((user_id)={id + 2})") # yyyy-mm-dd

let rst = ado.query($&"SELECT * FROM {tbl}")

for rowIdx, row in rst:
   for idx, name, fld in row.fields:
      echo "Row ", rowIdx, ", Fld[", idx," | ", name, "] = ", nz(fld, "!!! Not set !!!")
