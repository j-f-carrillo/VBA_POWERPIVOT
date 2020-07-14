# VBA_POWERPIVOT
VBA Macros to Add a Table of Macros to Data Model
Based Around 
With ActiveWorkbook.Model
        
        .ModelMeasures.Add mname, mTable, mformula, .ModelFormatDecimalNumber
        
    End With
    
    
    Measures Loop a Table on Active Sheet with Structure as Follows 
      Col 1         Col 2         Col 3
    Table_Name    Measure_Name   Formula_Str
