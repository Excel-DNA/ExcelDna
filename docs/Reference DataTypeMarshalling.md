---
layout: page
title: Excel - DNA Reference: Data type marshaling
---

The allowed function parameter and return types are:
* Double
* String
* DateTime    -- returns a double to Excel (maybe string is better to return?)
* Double[]()    -- if only one column is passed in, takes that column, else first row is taken 
* {"Double[,](,)"}
* Object
* Object[]()    -- if only one column is passed in, takes that column, else first row is taken 
* {"Object[,](,)"}
* Boolean (bool) -- returns an Excel bool (maybe string is better to return to Excel?)
* Int32 (int)
* Int16 (short)
* UInt16 (ushort)
* Decimal
* Int64 (long)

incoming function parameters of type Object will only arrive as one of the following:
* Double
* String
* Boolean
* ExcelDna.Integration.ExcelError
* ExcelDna.Integration.ExcelMissing
* ExcelDna.Integration.ExcelEmpty
* {"Object[,](,)"} containing an array with a mixture of the above types
* ExcelReference -- (Only if AllowReference=true in ExcelArgumentAttribute causing R type instead of P)

function parameters of type Object[]() or {"Object[,](,)(,)"} will receive an array containing a mixture of the above types (excluding {"Object[,](,)(,)"})

return values of type Object are allowed to be:
* Double
* String
* DateTime
* Boolean
* Double[]()
* {"Double[,](,)"}
* Object[]()
* {"Object[,](,)"}
* ExcelDna.Integration.ExcelError
* ExcelDna.Integration.ExcelMissing.Value // Converted by Excel to be 0.0
* ExcelDna.Integration.ExcelEmpty.Value   // Converted by Excel to be 0.0
* ExcelDna.Integration.ExcelReference
* Int32 (int)
* Int16 (short)
* UInt16 (ushort)
* Decimal
* Int64 (long)
otherwise return #VALUE! error

return values of type Object[]() and {"Object[,](,)"} are processed as arrays of the type Object, containing a mixture of the above, excluding the array types.

### Special return values

The following invalid return values, are returned to Excel as indicated
*    Object{"[0](0)"} => #VALUE
*    Object{"[0.0](0.0)"} => #VALUE
*    String.Empty => “”
*    ExcelEmpty.Value => 0
*    null => #NUM

