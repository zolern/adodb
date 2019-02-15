import adodb, random, strformat

let 
   connect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb"
   db = newADODB(connect)
   tbl = "users"
   uid = "user_id"


var userId: int = nz(db.dMax(uid, tbl), 0)

db.exec(sql"DELETE * FROM {tbl}")

proc iteration(id: var int) =
   id.inc

   db.exec(sql"INSERT INTO {tbl} ( {uid}, user_name, user_balance ) VALUES ({id}, 'NEW TEST USER {id}', {rand(1000.0)})")

   let 
      count = db.dCount(uid, tbl)
      sum: float = db.dSum("user_balance", tbl)
      avg: float = db.dAvg("user_balance", tbl)

   if count >= 1000:
      # keep count of records to not exceed 1000
      let minID = db.dMin(uid, tbl)
      db.exec(sql"DELETE * FROM {tbl} WHERE (({uid})={minID})")

   stdout.write('\r', fmt"MaxID: {id:>5}, Count: {count:>4}, Total sum: {sum:>9.2f}, Average: {avg:>6.2f}")

while true:
   iteration(userId)