import winim/com as wincom
import json, times, jsontime

proc newJDbTime*(dbTime: float): JsonNode = 
   let 
      begTm = toTime(DateTime(year: 2016, month: mJan, monthday: 1))
      elapsed = toInt((dbTime - 42370.0) * 24.0 * 60.0 * 60.0)
      dateTime: DateTime = utc(begTm + elapsed.seconds)

   return newJDateTime(dateTime)

proc newJTime*(v: variant): JsonNode =
   return newJDbTime(fromVariant[float](v))

proc variantToJson*(v: variant): JsonNode =
   case v.vt
   of 0: newJNull() # "VT_EMPTY"
   of 1: newJNull() # "VT_NULL"
   of 2: newJInt(fromVariant[int](v)) # "VT_I2"
   of 3: newJInt(fromVariant[int](v)) # ""VT_I4"
   of 4: newJFloat(fromVariant[float](v)) # "VT_R4"
   of 5: newJFloat(fromVariant[float](v)) # "VT_R8"
   of 6: newJNull() # "VT_CY"
   of 7: newJTime(v) # "VT_DATE"
   of 8: newJString(fromVariant[string](v)) # "VT_BSTR"
   of 9: newJNull() # "VT_DISPATCH"
   of 10: newJNull() # "VT_ERROR"
   of 11: newJBool(fromVariant[bool](v)) # "VT_BOOL"
   of 12: newJNull() # "VT_VARIANT"
   of 13: newJNull() # "VT_UNKNOWN"
   of 14: newJInt(fromVariant[int](v)) # "VT_DECIMAL"
   of 16: newJInt(fromVariant[int](v)) # "VT_I1"
   of 17: newJInt(fromVariant[int](v)) # "VT_UI1"
   of 18: newJInt(fromVariant[int](v)) # "VT_UI2"
   of 19: newJInt(fromVariant[int](v)) # "VT_UI4"
   of 20: newJInt(fromVariant[int](v)) # "VT_I8"
   of 21: newJInt(fromVariant[int](v)) # "VT_UI8"
   of 22: newJInt(fromVariant[int](v)) # "VT_INT"
   of 23: newJInt(fromVariant[int](v)) # "VT_UINT"
   of 24: newJNull() # "VT_VOID"
   of 25: newJNull() # "VT_HRESULT"
   of 26: newJNull() # "VT_PTR"
   of 27: newJNull() # "VT_SAFEARRAY"
   of 28: newJNull() # "VT_CARRAY"
   of 29: newJNull() # "VT_USERDEFINED"
   of 30: newJString(fromVariant[string](v)) # "VT_LPSTR"
   of 31: newJString(fromVariant[string](v)) # "VT_LPWSTR"
   of 36: newJNull() # "VT_RECORD"
   of 37: newJNull() # "VT_INT_PTR"
   of 38: newJNull() # "VT_UINT_PTR"
   of 64: newJTime(v) #"VT_FILETIME"
   of 65: newJNull() # "VT_BLOB"
   of 66: newJNull() # "VT_STREAM"
   of 67: newJNull() # "VT_STORAGE"
   of 68: newJNull() # "VT_STREAMED_OBJECT"
   of 69: newJNull() # "VT_STORED_OBJECT"
   of 70: newJNull() # "VT_BLOB_OBJECT"
   of 71: newJNull() # "VT_CF"
   of 72: newJNull() # "VT_CLSID"
   of 0xfff: newJNull() # "VT_BSTR_BLOB"
   else: newJNull() # "VT_ILLEGAL"
