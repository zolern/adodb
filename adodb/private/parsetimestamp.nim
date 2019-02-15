#====================================================================
#
#               adodb - Microsoft ADO DB accessor
#                (c) Copyright 2019 Encho "Zolern" Topalov
# 
#              SQL format, timestamp parser
#
#====================================================================

## SQL format parser helper for timestamp literal parsing
## #dd.mm.yyyy# is parsed as #mm/dd/yyyy#
## #yyyy-mm-dd# is parsed as #mm/dd/yyyy#
## 
## Mailformed literal is returns as is

import unicode

type
   ParserPhase = enum
      ppNone = 0, ppDate, ppTime, ppTimeEnd, ppSym1, ppSym2, ppFinal

   TimeStampParser* = ref object of RootObj
      elem: seq[string]
      wasDigit: bool
      phase: ParserPhase
      added: string
      dateDivider: string
      add12hours: bool
      dateParts: int
      isFailed: bool

proc newTimeStampParser*: TimeStampParser {.inline.} =
   ## Constructor
   new(result)
   result.added = "#"

func isOk*(self: TimeStampParser): bool = self.phase == ppFinal
func text*(self: TimeStampParser): string = self.added

proc parse(self: TimeStampParser): string =
   var date = ""
   
   if self.dateParts > 0:
      case self.dateDivider
      of ".":
         date &= $self.elem[1] & "/" & $self.elem[0] & "/" & $self.elem[2]
      of "-":
         date &= $self.elem[1] & "/" & $self.elem[2] & "/" & $self.elem[0]
      else:
         date &= $self.elem[0] & "/" & $self.elem[1] & "/" & $self.elem[2]
   
   let 
      timeParts = self.elem.len - self.dateParts 
      timeStarts = self.dateParts

   if timeParts > 0:
      date &= (if date != "": " " else: "") & self.elem[timeStarts]

   if timeParts > 1:
      date &= ":" & self.elem[timeStarts + 1]
   
   if timeParts > 2:
      date &= ":" & self.elem[timeStarts + 2]

   if self.add12hours:
      date &= " PM"    

   return "#" & date & "#"

proc addSymbol*(self: TimeStampParser; sym: string): bool {.discardable.} =
   ## acts as iterator through main symbol sequence started with #
   self.added &= sym

   if self.isFailed: return false

   self.isFailed = true

   case sym
   of "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
      if self.phase >= ppTimeEnd: return

      if not self.wasDigit:
         self.elem.add(sym)
      else:
         self.elem[^1] &= sym
      self.wasDigit = true

      if self.elem[^1].len > 4:
         return
   
   of " ", "/", "-", ".", ":":
      if self.phase >= ppSym1:
         if sym == ".":
            self.added = self.added[0..^2]
            self.isFailed = false
            return true
         else:
            return

      if not self.wasDigit: return
      
      self.wasDigit = false

      if sym == " ":
         if self.phase > ppTime: return
         if self.phase == ppDate:
            self.dateParts.inc
            self.phase = ppTime
         else:
            self.phase = ppTimeEnd
      elif sym == ":":
         if self.phase == ppNone: self.phase = ppTime
         if self.phase != ppTime: return
      else:
         if self.phase == ppNone: self.phase = ppDate
         if self.phase != ppDate: return

         if self.dateDivider == "": self.dateDivider = sym
         if self.dateDivider != sym: return

         self.dateParts.inc

   of "A", "a", "P", "p":
      if self.phase >= ppSym1: return
      if self.phase == ppDate: return
      if not (self.wasDigit or self.phase == ppTimeEnd): return

      self.phase = ppSym1
      if sym == "P" or sym == "p": self.add12hours = true
   
   of "M", "m":
      if self.phase != ppSym1: return
      self.phase = ppSym2
   
   of "#":
      if self.phase == ppSym1: return
      if self.phase != ppSym2:
         if self.wasDigit:
            if self.phase == ppDate: self.dateParts.inc
         else:
            return
      
      if self.dateParts > 3: return
      if self.dateParts == 2: return
      if self.elem.len - self.dateParts > 3: return
      
      self.phase = ppFinal
      self.added = self.parse
      return
   
   else:
      return
   
   self.isFailed = false
   return true

when isMainModule:
   proc parseDate(date: string): (string, bool) =
      var tsp = newTimeStampParser()
      var fullCheck = false
      var firstSym = true
      var sym: string
      var isOK: bool

      if date == "":
         tsp.addSymbol("#")
         tsp.addSymbol("#")
      else:
         for r in date.runes:
            sym = $r
            
            if firstSym:
               firstSym = false
               if sym == "#":
                  fullCheck = true
               else:
                  tsp.addSymbol("#")

            tsp.addSymbol(sym)
      
         if not fullCheck:
            tsp.addSymbol("#")

      if not fullCheck:
         return (tsp.text[1..^2], tsp.isOk)
      else:
         return (tsp.text, tsp.isOk)

   proc check(test, expectedText: string): bool {.discardable.} =
      let (res, isOk) = parseDate(test)
      if res != expectedText:
         echo "Expected '" & test & "' to produce '" & expectedText & "', but received '" & res & "'"
         quit(-1)

      return isOk
      
   proc check(test, expectedText: string, expectedOk: bool) =
      let isOk = check(test, expectedText)
      if isOk != expectedOk:
         echo "Expected '" & test & "' to produce '" & $expectedOk & "', but received '" & $isOk & "'"
         quit(-1)

   check("", "", false)
   check(" ", " ", false)
   check("#", "#", false)
   check("##", "##", false)
   check("###", "###", false)
   check("test", "test", false)
   check("pm", "pm", false)
   check("1", "1", true)
   check("12", "12", true)
   check("123", "123", true)
   check("1234", "1234", true)
   check("12345", "12345", false)
   check("1am", "1", true)
   check("1 am", "1", true)
   check("1pm", "1 PM", true)
   check("1 pm", "1 PM", true)
   check("1 p.m.", "1 PM", true)
   check("1 p.m", "1 PM", true)
   check("1 pm.", "1 PM", true)
   check("1:2", "1:2", true)
   check("1:2:3", "1:2:3", true)
   check("1:2:3:4", "1:2:3:4", false)
   check("#1 2#", "#1 2#", false)
   check("1.2", "1.2", false)
   check("1.2.3", "2/1/3", true)
   check("1/2", "1/2", false)
   check("1/2/3", "1/2/3", true)
   check("1/2/3/4", "1/2/3/4", false)
   check("1-2", "1-2", false)
   check("1-2-3", "2/3/1", true)
   check("1.2/3", "1.2/3", false)
   check("1/2-3", "1/2-3", false)
   check("1/2/3 ", "1/2/3 ", false)
   check("1/2/3pm", "1/2/3pm", false)
   check("1/2/3 pm", "1/2/3 pm", false)
   check("1/2/3 test", "1/2/3 test", false)
   check("1/2/3 4", "1/2/3 4", true)
   check("1/2/3 4:5", "1/2/3 4:5", true)
   check("1/2/3 4:5am", "1/2/3 4:5", true)
   check("1/2/3 4:5 am", "1/2/3 4:5", true)
   check("1/2/3 4:5pm", "1/2/3 4:5 PM", true)
   check("1/2/3 4:5:6", "1/2/3 4:5:6", true)
   check("1/2/3 4:5:6 7", "1/2/3 4:5:6 7", false)
   check("1/2/3 4:5:6am", "1/2/3 4:5:6", true)
   check("1/2/3 4:5:6 am", "1/2/3 4:5:6", true)
   check("1/2/3 4:5:6pm", "1/2/3 4:5:6 PM", true)
   check("1/2/3 4:5:6 pM", "1/2/3 4:5:6 PM", true)
