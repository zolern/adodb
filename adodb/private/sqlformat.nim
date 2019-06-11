#====================================================================
#
#               adodb - Microsoft ADO DB accessor
#                (c) Copyright 2019 Encho "Zolern" Topalov
# 
#                   SQL format & literal interpolation
#
#====================================================================

##[ SQL string literals interpolation with `{fields}` and `#timestamps#` inputs

### Fields input:

It is similar to strformat string interpolation. Symbols `{` and `}`
should be excaped with doubling it: `{{` and `}}`

.. code-block:: nim

    let userName = "Test User"
    echo sql"SELECT * FROM Users WHERE ((user_name)='{userName}')"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_name)='Test User')


### Field type

It is possible to set "type" of field - string or timestamp, so appropriate value will be
correct delimited with single quotes (i.e. as 'text field') for string and with ## 
(i.e. as #timestamp#) for timestamps. Field should be prefixed with $ for string field type 
and with # for timestamp field type. Symbols $ and # should be doubled for excaping

.. code-block:: nim
    let findUser = "Test user 1"
    let bornDate = initDateTime(27, tm.Month(3), 1954, 0, 0, 0)
    echo sql"SELECT * FROM Users WHERE ((user_name)=${findUser}) AND ((user_bday)=>#{bornDate})"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_name)='Test user 1') AND ((user_bday)=>#3/27/1954#)

### Special timestamp construction

Special construction form for timestamps: ```#{ <year>, <month>, <day> [, <hour> [, <minutes> [, <seconds>]]] }```

.. code-block:: nim
    let findUser = "Test user 1"
    
    echo sql"SELECT * FROM Users WHERE ((user_bday)=>#{1999, 12, 31})"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_bday)=>#12/31/1999#)

### Timestamp literals and custom formats.

Timestamp literal should be delimited by ##.
   
.. code-block:: nim

    echo sql"SELECT * FROM Users WHERE ((user_bday)=>#12/31/1999#)"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_bday)=>#12/31/1999#)

#### -- Date in timestamp literal

Regardless that SQL standard requires date in timestamps to be in form 
`m(onth)/d(ay)/y(ear)` it is possible to set regional specific (custom) formats:

     * `#d.m.y#` (with dot as field delimiter)
   
     * `#y-m-d#` (with dash as field delimiter)

.. code-block:: nim

    echo sql"SELECT * FROM Users WHERE ((user_bday)=>#31.12.1999#)"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_bday)=>#12/31/1999#)

    echo sql"SELECT * FROM Users WHERE ((user_bday)=>#1999-12-31#)"
    
    # Output:
    # SELECT * FROM Users WHERE ((user_bday)=>#12/31/1999#)

#### -- Time in timestamp literal

Time in timestamp literal is optional. If set it should be in form 
`h(our):m(inutes):s(econds)` with optional AM / PM suffixes

.. code-block:: nim

    echo sql"UPDATE operations SET created=#15.1.1998 14:03#"
    
    # Output:
    # UPDATE operations SET created=#15/1/1998 14:03#

    echo sql"UPDATE operations SET created=#15.1.1998 2:03p.m.#"
    
    # Output:
    # UPDATE operations SET created=#15/1/1998 2:03 PM#

]##

import macros, parseutils, unicode
import strutils
import parsetimestamp, parsefield

from times as tm import nil

template dbFormatTS*(dt: tm.DateTime): string =
   ##[ format DateTime as SQL compatible string

   .. code-block:: nim
       let bornDate = initDateTime(27, tm.Month(3), 1954, 0, 0, 0)

       echo sql"SELECT * FROM Users WHERE ((user_bday)=#{bornDate})"

       # Output:
       # SELECT * FROM Users WHERE ((user_bday)=#3/27/1954#)
   ]##
   var formatTS = "MM/dd/yyyy"
   if dt.hour != 0 or dt.minute != 0 or dt.second != 0: 
      formatTS.add " HH:mm:ss"
   
   tm.format(dt, formatTS)

template dbFormatTS*(year, month, day: int; hour: int = 0; minute: int = 0; second: int = 0): string =
   ##[ create timestamp from date/time components and format it for SQL
   
   .. code-block:: nim
       echo sql"SELECT * FROM Users WHERE ((user_bday)=#{1954, 3, 27})"

       # Output:
       # SELECT * FROM Users WHERE ((user_bday)=#3/27/1954#)
   ]##
   dbFormatTS(tm.initDateTime(day, tm.Month(month), year, hour, minute, second))

template callField(res, arg) {.dirty.} =
   when arg is string:
      res.add arg
   else:
      res.add($arg)

macro `$&`*(pattern: string): string =
   type TermType = enum ttNone, ttSpecField, ttTimeStamp, ttOpenBracket, ttInputField, ttCloseBracket
   type TermSpecField = enum sfNone = (0, "None"), sfTimeStamp = "TS", sfString = "Str"

   ## For a specification of the ``$&`` macro, see the module level documentation.
   if pattern.kind notin {nnkStrLit..nnkTripleStrLit}:
      error "sql formatting (sql(), &) only works with string literals",
         pattern
   let 
      f = pattern.strVal
      res = genSym(nskVar, "fmtSqlRes")

   result = newNimNode(nnkStmtListExpr, lineInfoFrom = pattern)
   result.add newVarStmt(res, newCall(bindSym"newStringOfCap",
         newLit(f.len + count(f, '#')*10 + count(f, '{')*10 )))

   var 
      strlit = ""
      term: TermType = ttNone
      specField: TermSpecField = sfNone
      tsp: TimeStampParser
      fp: FieldParser
   
   for r in f.runes:
      let sym = $r
      
      case term
      of ttNone:
         case sym
         of "$":
            term = ttSpecField
            specField = sfString

         of "#":
            term = ttSpecField
            specField = sfTimeStamp
         
         of "{":
            term = ttOpenBracket
            fp = newFieldParser()
         
         of "}":
            term = ttCloseBracket

         else:
            strlit.add sym
      
      of ttSpecField:
         if sym == "{":
            term = ttOpenBracket
            fp = newFieldParser()

         else:
            let tempSpecField = specField
            specField = sfNone

            case tempSpecField
            of sfTimeStamp:
               if sym == "#":
                  term = ttNone
                  strlit.add "#"
               else:      
                  term = ttTimeStamp
                  tsp = newTimeStampParser()
                  
                  if not tsp.addSymbol(sym):
                     strlit.add tsp.text
                     term = ttNone
                     tsp = nil
      
            of sfString:
               term = ttNone
               strlit.add "$"
               if sym != "$":
                  strlit.add sym

            else:
               term = ttNone
               strlit.add sym
         
      of ttTimeStamp:
         if not tsp.addSymbol(sym):
            strlit.add tsp.text
            term = ttNone
            tsp = nil
      
      of ttOpenBracket:
         if not fp.addSymbol(sym):
            strlit.add fp.text
            term = ttNone
            fp = nil
         else:
            strLit.add case specField 
               of sfString: "'" 
               of sfTimeStamp: "#" 
               else: ""
            if strlit.len > 0:
               result.add newCall(bindSym"add", res, newLit(strlit))
               strlit = ""         
            term = ttInputField

      of ttInputField:
         if not fp.addSymbol(sym):
            let x = parseExpr(case specField
               of sfTimeStamp:
                  "dbFormatTS(" & fp.text & ")"
               else:
                  fp.text
            )
                  
            result.add getAst(callField(res, x))
            term = ttNone
            
            strLit.add case specField 
               of sfString: "'" 
               of sfTimeStamp: "#"
               else: ""
            
            specField = sfNone
            fp = nil
      
      of ttCloseBracket:
         if sym == "}":
            strlit.add "}"
            term = ttNone
         else:
            error "invalid format string: '}' instead of '}}'", pattern
         
   case term
   of ttNone:
      discard
   of ttSpecField:
      strlit.add case specField
         of sfTimeStamp:
            "#"
         of sfString:
            "$"
         else:
            ""
   of ttTimeStamp:
      strlit.add tsp.text
   of ttOpenBracket:
      error "invalid format string: '{' instead of '{{'", pattern
   of ttInputField:
      error "invalid format string: missing '}'", pattern
   of ttCloseBracket:
      error "invalid format string: '}' instead of '}}'", pattern
   
   if strlit.len > 0:
      result.add newCall(bindSym"add", res, newLit(strlit))
   
   result.add res

template sql*(pattern: string): untyped =
   ## An alias for ``$&``.
   bind `$&`
   $&pattern

when isMainModule:
   template check(test, expected: string): untyped =
      let res = $&test
      if res != expected:
         echo "Expected '" & test & "' to produce '" & expected & "', but received '" & res & "'"
         quit(-1)
      
   check("", "")
   check("no fields or timestamps", "no fields or timestamps")         
   check("Escaped open bracked: '{{'", "Escaped open bracked: '{'")
   check("Escaped open bracked: '}}'", "Escaped open bracked: '}'")

   let s = "string"
   check("test with '{s}' field", "test with 'string' field")
   check("Field with brackets: {{{s}}}", "Field with brackets: {string}")

   check("#", "#")
   check("$", "$")
   check("##", "#")
   check("$$", "$")
   check("$#", "$#")
   check("#$", "#$")

   check("Hello #ABC", "Hello #ABC")
   check("#123", "#123")
   check("#ABC", "#ABC")
   check("$123", "$123")
   check("Hello #ABC #123 #$test $#TEST $123 $${{}}##{{}}", "Hello #ABC #123 #$test $#TEST $123 ${}#{}")

   # dbFormatS
   check("${s}", "'string'")

   # dbFormatTS
   let dt = tm.initDateTime(23, tm.Month(11), 1988, 10, 0, 0)
   check("Test #{dt}", "Test #11/23/1988 10:00:00#")
   check("Test #{1999, 12, 31}", "Test #12/31/1999#")

   check("compile time #9.2.2019 15:45#", "compile time #2/9/2019 15:45#")
   check("compile time #1999-7-12 10:01pm#", "compile time #7/12/1999 10:01 PM#")
   
   let idx = 123
   check("${idx}", "'123'")
   check("$${idx}", "$123")
   check("##{idx}", "#123")