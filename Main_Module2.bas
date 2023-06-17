Attribute VB_Name = "Main"
Sub Market_Analysis()

Dim ws As Worksheet
'Full Clear
For Each ws In ThisWorkbook.Worksheets
    ws.Columns("H:BB").Delete
Next ws


Dim StartTime As Double
StartTime = Timer
Dim EndTime As Double
Dim TotalTime As Double
Dim FirstBlock As Double
Dim SecondBlock As Double
Dim ThridBlock As Double
Dim FourthBlock As Double

CallColors


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK START - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Dim the needed dictionaries
Dim Ticker As New Scripting.Dictionary
Dim TickerKey As Variant
Dim WorksheetData As New Scripting.Dictionary
Dim WorksheetName As Variant 'Key


'Dim the temporary Looping variables
Dim rng As Range
Dim cell As Range
Dim VolSum As Double
Dim Vol As Double
Dim Op As Double
Dim Cl As Double
Dim Change As Double
Dim Earliest_Date As Date
Dim Latest_Date As Date
Dim First_Open As Double
Dim Yr_Chg As Double
Dim Pct_Chg As Double
Dim Last_Close As Double
Dim Largest_Increase_T As String
Dim Largest_Increase_V As Double
Dim Largest_Decrease_T As String
Dim Largest_Decrease_V As Double
Dim Largest_Volume_T As String
Dim Largest_Volume_V As Double
Dim Off As Double


Dim dt As Date
Dim dtArray As Variant




Dim LastRow As Double
Dim LastColumn As Double
Dim SafeWorkSpace As String
Dim SafeWorkHeaders As String


FirstBlock = Timer - StartTime
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK END - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -



'Count the rows and columsn to establish the range, store those in a value
'WorksheetData(WorkSheetName)(LastRow (0), LastColumn(1), SafeWorkSpace(2), SafeWorkHeaders(3))
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK START - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
For Each ws In ThisWorkbook.Worksheets
    WorksheetName = ws.Name
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    SafeWorkSpace = ws.Cells(2, LastColumn + 3).Address
    SafeWorkHeaders = ws.Cells(1, LastColumn + 3).Address
    WorksheetData.Add ws.Name, Array(LastRow, LastColumn, SafeWorkSpace, SafeWorkHeaders)
    
Next ws
SecondBlock = Timer - StartTime - FirstBlock
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK END - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'=====================================================================NOTE/TEST BLOCK======================================================================
'============Test Outputs
'For Each WorksheetName In WorksheetData.Keys
'    MsgBox ("Worksheet Name: " & WorksheetName & ", Last Row: " & WorksheetData(WorksheetName)(0) & ", Last Column: " & WorksheetData(WorksheetName)(1))
'Next WorksheetName

'MsgBox (WorksheetData.Item(ActiveSheet.Name))
'=====================================================================NOTE/TEST BLOCK======================================================================

'Prepare the Worksheets

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK START - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
For Each ws In ThisWorkbook.Worksheets
    ws.Range("H1").Value = "<DateCorr>"
    
    'Assign the new table and row areas as X number of columns away from the data set
    ws.Range(WorksheetData(ws.Name)(3)).Value = "Ticker"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(0, 1).Value = "Yearly Change"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(0, 2).Value = "Percent Change"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(0, 3).Value = "Total Stock Volume"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(0, 7).Value = "Ticker"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(0, 8).Value = "Value"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(1, 6).Value = "Greatest % Increase"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(2, 6).Value = "Greatest % Decrease"
    ws.Range(WorksheetData(ws.Name)(3)).Offset(3, 6).Value = "Greatest Total Volume"
    
    'For Each cell In ws.Range("H2:H" & WorksheetData.Item(ws.Name)(0))
    'dt = DateSerial(Left(cell.Offset(0, -6).Value, 4), Mid(cell.Offset(0, -6), 5, 2), Right(cell.Offset(0, -6), 2))
    'cell.Value = dt
    'Next cell
    
    
    dtArray = ws.Range("B2:B" & WorksheetData.Item(ws.Name)(0)).Value
    
    Dim i As Long
    For i = LBound(dtArray, 1) To UBound(dtArray, 1)
        dt = DateSerial(Left(dtArray(i, 1), 4), Mid(dtArray(i, 1), 5, 2), Right(dtArray(i, 1), 2))
        dtArray(i, 1) = dt
    Next i
    
    ws.Range("H2:H" & WorksheetData.Item(ws.Name)(0)).Value = dtArray
    
    ws.Columns.AutoFit
    ws.Columns.HorizontalAlignment = xlCenter
    
Next ws
ThirdBlock = Timer - StartTime - FirstBlock - SecondBlock
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK END - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Populate the Ticker Dictionary Ticker(TickerKey)(Earliest_Date(0), First_Open(1), Latest_Date(2), Last_Close(3), VolSum(4), Pct_Chg(5), Yr_Chg(6))
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK START - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
For Each ws In ThisWorkbook.Worksheets

'Reset Variables Per ws
    'For Each TickerKey In Ticker.Keys
        'Ticker(TickerKey) = Array(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        'Ticker(TickerKey) = Array(Earliest_Date, First_Open, Latest_Date, Last_Close, VolSum, Pct_Chg, Yr_Chg)
    'Next TickerKey
Set Ticker = CreateObject("Scripting.Dictionary")
Set Ticker = Nothing

Debug.Print "Ticker Count is: " & Ticker.Count&; ", Now Processing WorkSheet " & ws.Name

First_Open = 0
Last_Close = 0
Earliest_Date = DateSerial(2050, 12, 25)
Latest_Date = DateSerial(1900, 1, 1)
VolSum = 0
Pct_Chg = 0
Yr_Chg = 0
Largest_Increase_T = "UNK"
Largest_Increase_V = -0.01
Largest_Decrease_T = "UNK"
Largest_Decrease_V = 0.01
Largest_Volume_T = "UNK"
Largest_Volume_V = -1
    'For Paste: Ticker(TickerKey) = Array(Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(6))
    'Analyze All of the Rows
    For Row = 2 To WorksheetData(ws.Name)(0)
        TickerKey = ws.Cells(Row, 1).Value
        Vol = ws.Cells(Row, 7).Value
        If Not Ticker.Exists(TickerKey) Then
            Ticker.Add TickerKey, Array(Earliest_Date, First_Open, Latest_Date, Last_Close, VolSum, Pct_Chg, Yr_Chg)
        Else
            Ticker(TickerKey) = Array(Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4) + Vol, Ticker(TickerKey)(5), Ticker(TickerKey)(6))
        End If
        
        'Debug.Print TickerKey, Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(6)

        dt = ws.Cells(Row, 8).Value
        
            If Ticker(TickerKey)(0) > dt Then
                'Debug.Print "Enter Earlier IF"
                Ticker(TickerKey) = Array(dt, ws.Cells(Row, 3).Value, Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(6))
            ElseIf dt > Ticker(TickerKey)(2) Then
                'Debug.Print "Enter Later IF"
                Ticker(TickerKey) = Array(Ticker(TickerKey)(0), Ticker(TickerKey)(1), dt, ws.Cells(Row, 6).Value, Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(6))
            End If
        'Debug.Print TickerKey, Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Cells(Row, 6).Value, Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(6)
        'Debug.Print "Date: " & dt
        'Debug.Print "Earliest Date: " & Ticker(TickerKey)(0)
        'Debug.Print "Latest Date: " & Ticker(TickerKey)(2)
    Next Row
    
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK END - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'REFERENCE Ticker(TickerKey)(Earliest_Date(0), First_Open(1), Latest_Date(2), Last_Close(3), VolSum(4), Pct_Chg(5), Yr_Chg(6))
    'Reset Offset Var
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK START - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Off = 0
    
    For Each TickerKey In Ticker.Keys
    If TickerKey <> "" Then
        'Update Tick Metrics and Format Cells
        'Debug.Print TickerKey
        'Debug.Print Ticker(TickerKey)(1)
        'Debug.Print ws.Name
        Ticker(TickerKey) = Array(Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4), Ticker(TickerKey)(5), Ticker(TickerKey)(3) - Ticker(TickerKey)(1))
        Ticker(TickerKey) = Array(Ticker(TickerKey)(0), Ticker(TickerKey)(1), Ticker(TickerKey)(2), Ticker(TickerKey)(3), Ticker(TickerKey)(4), Ticker(TickerKey)(6) / Ticker(TickerKey)(1), Ticker(TickerKey)(6))
        ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 0).Value = TickerKey
        ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 1).Value = Round(Ticker(TickerKey)(6), 2)
        If Ticker(TickerKey)(6) < 0 Then
            With ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 1)
                .Interior.Color = Red
                .Font.Color = White
            End With
        ElseIf Ticker(TickerKey)(6) > 0 Then
            ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 1).Interior.Color = Green
        ElseIf Ticker(TickerKey)(6) = 0 Then
            ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 1).Interior.Color = LightBlue
        End If
        
        
        With ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 2)
            .Value = Ticker(TickerKey)(5)
            .NumberFormat = "0.00%"
        End With
        If Ticker(TickerKey)(5) < 0 Then
            With ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 2)
                .Interior.Color = Red
                .Font.Color = White
            End With
        ElseIf Ticker(TickerKey)(5) > 0 Then
            ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 2).Interior.Color = Green
        ElseIf Ticker(TickerKey)(5) = 0 Then
            ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 2).Interior.Color = LightBlue
        End If
        ws.Range(WorksheetData(ws.Name)(2)).Offset(Off, 3).Value = Ticker(TickerKey)(4)
        
        'Check for greatest value in arrays during the ticker loop
        
            If Ticker(TickerKey)(5) > Largest_Increase_V Then
                Largest_Increase_V = Ticker(TickerKey)(5)
                Largest_Increase_T = TickerKey
            ElseIf Ticker(TickerKey)(5) < Largest_Decrease_V Then
                Largest_Decrease_V = Ticker(TickerKey)(5)
                Largest_Decrease_T = TickerKey
            End If
        
            If Ticker(TickerKey)(4) > Largest_Volume_V Then
                Largest_Volume_V = Ticker(TickerKey)(4)
                Largest_Volume_T = TickerKey
            End If
    
        Off = Off + 1
     End If
        
    Next TickerKey
    
    ws.Range(WorksheetData(ws.Name)(2)).Offset(0, 7).Value = Largest_Increase_T
        With ws.Range(WorksheetData(ws.Name)(2)).Offset(0, 8)
            .Value = Largest_Increase_V
            .NumberFormat = "0.00%"
        End With
    ws.Range(WorksheetData(ws.Name)(2)).Offset(1, 7).Value = Largest_Decrease_T
        With ws.Range(WorksheetData(ws.Name)(2)).Offset(1, 8)
            .Value = Largest_Decrease_V
            .NumberFormat = "0.00%"
        End With
    ws.Range(WorksheetData(ws.Name)(2)).Offset(2, 7).Value = Largest_Volume_T
        With ws.Range(WorksheetData(ws.Name)(2)).Offset(2, 8)
            .Value = Largest_Volume_V
            .NumberFormat = "0.00E+00"
        End With
        
    ws.Columns.AutoFit
    ws.Columns.HorizontalAlignment = xlCenter

Next ws
FourthBlock = Timer - StartTime - FirstBlock - SecondBlock - ThirdBlock
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - CODE BLOCK END - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Perform arithmetic on stored values to create output values. Stored values into variables

'Update Range values (either per sheet or loop through sheets)

'Conditionally format values / ranges

EndTime = Timer
TotalTime = EndTime - StartTime

Debug.Print "Time taken for First Block: " & FirstBlock & " seconds"
Debug.Print "Time taken for Second Block: " & SecondBlock & " seconds"
Debug.Print "Time taken for Third Block: " & ThirdBlock & " seconds"
Debug.Print "Time taken for Fourth Block: " & FourthBlock & " seconds"
Debug.Print "Total run time is: " & TotalTime & " seconds"
End Sub

Sub Clear()
 For Each ws In ThisWorkbook.Worksheets
    ws.Range("H:BB").Clear
    ws.Range("H:BB").ClearFormats
 Next ws
End Sub
