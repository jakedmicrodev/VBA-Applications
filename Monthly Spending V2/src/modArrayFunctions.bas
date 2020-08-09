Attribute VB_Name = "modArrayFunctions"
Option Explicit

Public Function SelectGroupsArray() As Variant

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GROUPS_WORKSHEET)

    Dim lastRow As Integer
    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row

    'https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    SelectGroupsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
    
cleanup:

    Set ws = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modArrayFunctions.SelectGroupsArray", Err.Description, 8
    GoTo cleanup
End Function

Public Function SelectHeadingEndsArray() As Variant

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HEADING_ENDS_WORKSHEET)
    
    Dim lastRow As Integer
    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).row
    
    Dim ends As Variant
    SelectHeadingEndsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
    
cleanup:

    Set ws = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modArrayFunctions.SelectHeadingEndsArray", Err.Description, 31
    GoTo cleanup
End Function

Public Function SelectMonthsArray() As Variant

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MONTHS_WORKSHEET)
        
    Dim lastRow As Integer
    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).row
    
    'https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    SelectMonthsArray = Application.Transpose(ws.Range("A1", "A" & lastRow + 1))
    
cleanup:

    Set ws = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modArrayFunctions.SelectMonthsArray", Err.Description, 54
    GoTo cleanup
End Function

Public Function SelectQueriesArray() As Variant

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(QUERIES_WORKSHEET)
    
    Dim lastRow As Integer
    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row

    'https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    SelectQueriesArray = Application.Transpose(ws.Range("A2", "B" & lastRow))
    
cleanup:

    Set ws = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modArrayFunctions.SelectQueriesArray", Err.Description, 77
    GoTo cleanup
End Function

Public Function SelectWorksheetsArray() As Variant

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WORKSHEETS_WORKSHEET)
    
    Dim lastRow As Integer
    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row

    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    SelectWorksheetsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
    
cleanup:

    Set ws = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modArrayFunctions.SelectWorksheetsArray", Err.Description, 100
    GoTo cleanup
End Function

