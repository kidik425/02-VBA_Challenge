Attribute VB_Name = "Module1"
Sub Main()

    Dim cSheet As String
    Dim cntSheets As Integer
    Dim i As Integer  'Iterator
    Dim s As Integer
    
    i = 1
    Worksheets(1).Activate
    
    'Get count of sheets before the consolidated.  This will be used as an iterator.
    s = Sheets.Count
    
    'Loop through worksheet create table
    For n = 1 To s
        'Assign current worksheet name to variable.  This will be passed into methods as needed
        cSheet = Sheets(i).Name
        
        'loop through each sheet and create analysis tables
        createVals (cSheet)
        
        evalVals (cSheet)
        
        i = i + 1
    Next n
        
    Worksheets(1).Activate
    
End Sub

Sub createVals(sht As String)

    'Declare object variables to hold references to worksheet containing cell range, and cell range itself
    Dim myWorksheet As Worksheet
    
    'The following Dims are value holders
    Dim tckr As String
    Dim yr As Integer
    Dim opn As Double
    Dim cls As Double
    Dim vol As Double
    
    tckr = ""
    yr = 0#
    opn = 0#
    cls = 0#
    vol = 0
    
    'Define iterators
    Dim c As Long 'for the Forloop
    Dim r As Integer 'for Output
    
    'Start of For Loop
    r = 2
    
    'Set focus to new worksheet
    Set myWorksheet = ThisWorkbook.Worksheets(sht)
    myWorksheet.Activate
    
    'Clear any previous results and Setup headers
    Range("I:Z").Clear
    Range("J1") = "Ticker"
    Range("K1") = "Year"
    Range("L1") = "Start"
    Range("M1") = "End"
    Range("N1") = "Yearly Change"
    Range("O1") = "Percent Change"
    Range("P1") = "Total Stock Volumne"
    
    c = Cells(Rows.Count, 1).End(xlUp).Row
 
    For n = 2 To c
    
    'Populate starting values
       
    If tckr <> Range("A" & n) And Range("C" & n) <> 0 Then
        'Get values
        tckr = Range("A" & n)
        yr = Left(Range("B" & n), 4)
        opn = Range("C" & n)
        
        'Populate cells with initial values
        Range("J" & r) = tckr
        Range("K" & r) = yr
        Range("L" & r) = opn
        
    End If
    
    'Populate iterative values
    cls = Range("F" & n)
    vol = vol + Range("G" & n)
    
    'Check next record for change of ticker
    If Range("A" & n) <> Range("A" & n + 1) Then
        Range("M" & r) = cls
        Range("P" & r) = vol
        
        'Reset vol to 0 for next record grouping
        vol = 0
        
        'Create and populate Year Change and Percent Change
        yc = cls - opn
        pc = yc / opn
        
        Range("N" & r) = yc
        Range("O" & r) = Round(pc, 4)
        Range("O" & r).NumberFormat = "0.00%"
        
        If pc < 0 Then
            Range("O" & r).Interior.Color = vbRed
        Else: Range("O" & r).Interior.Color = vbGreen
        End If
        
        'Iterate R
        r = r + 1
    End If
    
   Next
    
    
End Sub

Sub evalVals(sht As String)

    'Declare object variables to hold references to worksheet containing cell range, and cell range itself
    Dim myWorksheet As Worksheet
    
    'The following Dims are value holders
    Dim tckr As Variant
    Dim mx As Double
    Dim mn As Double
    Dim mv As Double
    Dim rngLookup As Range
    tckr = ""

    'Start of For Loop
    r = 2
    
    'Set focus to new worksheet
    Set myWorksheet = ThisWorkbook.Worksheets(sht)
    myWorksheet.Activate
    Set rngLookup = myWorksheet.Columns("I:O")
    
    'Create headers
    Range("S1") = "Ticker"
    Range("T1") = "Value"
    Range("R2") = "Greatest % Increase"
    Range("R3") = "Greatest % Decrease"
    Range("R4") = "Greatest Total Volume"
    Range("R:R").Columns.AutoFit
    
    'Find Values
    mx = WorksheetFunction.max(Range("O:O")) 'Biggest incr
    mn = WorksheetFunction.Min(Range("O:O")) ' Smallest incr
    mv = WorksheetFunction.max(Range("P:P")) ' Most vol

    'Populate and format values
    Range("T2") = Round(mx, 4)
    Range("T2").NumberFormat = "0.00%"
    Range("T3") = Round(mn, 4)
    Range("T3").NumberFormat = "0.00%"
    Range("T4") = mv
    
    Range("S2") = getTckr(mx, "O")
    Range("S3") = getTckr(mn, "O")
    Range("S4") = getTckr(mv, "P")
    
End Sub

Function getTckr(val As Double, rng As String) As String

    Dim tckr As String
    
    'Define iterators
    Dim c As Long 'for the Forloop
    c = Cells(Rows.Count, 1).End(xlUp).Row
 
    For n = 2 To c
       
    If val = Range(rng & n) Then
        'Return Ticker
        tckr = Range("J" & n)
        Exit For
    End If
    
    Next n
    
    getTckr = tckr
End Function
