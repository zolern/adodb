import times, json

proc newJDateTime*(dt: DateTime): JsonNode =
   return %* {
         "year": dt.year, 
         "month": ord(dt.month),
         "day": ord(dt.monthday),
         "hour": ord(dt.hour),
         "minutes": ord(dt.minute),
         "seconds": ord(dt.second)
      }

proc getDateTime*(n: JsonNode): DateTime =
   ## Retrieves the DateTime value of a `JObject JsonNode`.
   if n.isNil or n.kind != JObject:
      return

   return DateTime(
      monthday: MonthdayRange(getOrDefault(n, "day").getInt(1)),
      year: getOrDefault(n, "year").getInt(0),
      month: Month(getOrDefault(n, "month").getInt(1)),
      hour: HourRange(getOrDefault(n, "hour").getInt(0)),
      minute: MinuteRange(getOrDefault(n, "minutes").getInt(0)),
      second: SecondRange(getOrDefault(n, "seconds").getInt(0))
   )
