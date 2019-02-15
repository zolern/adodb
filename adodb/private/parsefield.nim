#====================================================================
#
#               adodb - Microsoft ADO DB accessor
#                (c) Copyright 2019 Encho "Zolern" Topalov
# 
#              SQL format, { field } parser
#
#====================================================================

## SQL format parser helper for field literal parsing
## { and } are escaped with doubling it ({{ for { and }} for })

import unicode

type
   ParserPhase = enum
      ppFirstOpen, ppField, ppFinal

   FieldParser* = ref object of RootObj
      field: string
      phase: ParserPhase
      isFailed: bool

proc newFieldParser*: FieldParser {.inline.} =
   ## Constructor
   new(result)
   result.phase = ppFirstOpen

proc isOk*(self: FieldParser): bool = self.phase == ppFinal
   ## FieldParser.text contains field

proc text*(self: FieldParser): string = self.field
   ## return parsed field

proc addSymbol*(self: FieldParser; sym: string): bool {.discardable.} =
   ## acts as iterator though symbol sequence, started by {

   case sym
   of "{":
      if self.phase == ppFirstOpen:
         self.phase = ppFinal
         self.field = "{"
         return false
      else:
         doAssert(false, "invalid sequence: unexpected '{'")
   
   of "}":
      if self.phase != ppFinal:
         self.phase = ppFinal
         return false
      else:
         doAssert(false, "invalid sequence: unexpected '}'")
   
   else:
      if self.phase == ppFirstOpen:
         self.phase = ppField
      
      if self.phase == ppField:
         self.field &= sym
      else:
         doAssert(false, "invalid format string: unexpected symbol")
      
   return true
   
when isMainModule:
   proc parseField(field: string): (string, bool) =
      var 
         fp = newFieldParser()
         sym: string
         firstIteration = true
      
      for r in field.runes:
         sym = $r
         
         if firstIteration:
            firstIteration = false
            if sym == "{":
               continue
            else:
               doAssert(false, "should start with '{'")

         fp.addSymbol(sym)
      
      return (fp.text, fp.isOk)

   proc check(test, expectedText: string; expectedOk: bool) =
      let (res, isOk) = parseField(test)
      if res != expectedText:
         echo "Expected '" & test & "' to produce '" & expectedText & "', but received '" & res & "'"
         quit(-1)
      
      if isOk != expectedOk:
         echo "Expected '" & test & "' to produce '" & $expectedOk & "', but received '" & $isOk & "'"
         quit(-1)
   
   proc checktry(test: string; expectedException: bool) =
      var withException = true
      try:
         discard parseField(test)
         withException = false
      except:
         discard
      finally:
         if withException != expectedException:
            echo "Expected '" & test & "' to " & (if expectedException: "" else: "NOT") &
               "cause exception"
            quit(-1)
         
   check("{", "", false)
   check("{}", "", true)
   check("{{", "{", true)
   check("{test}", "test", true)
   check("{test", "test", false)
   checktry("{test{test2", true)
