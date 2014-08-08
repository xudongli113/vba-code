Attribute VB_Name = "Ä£¿é1"
Sub MacroDATA1()
'
' MacroDATA1 Macro
'
 Dim arr(36) As String

    Dim i As Integer
    i = 0
   For Each c In Worksheets("Sheet1").Range("A2:A37")

       arr(i) = Chr(34) & c.Value & Chr(34)

        i = i + 1

    Next c

Dim step As Integer
step = 3

Sheets("market_indices").Select
For i = 0 To 35
    Cells(1, i * step + 1).Value = Worksheets("Sheet1").Cells(i + 2, 1)
    Cells(2, i * step + 1).Select
    ActiveCell.FormulaR1C1 = _
        "=TR(" & arr(i) & ",""TR.IndexConstituentRIC;TR.IndexConstituentName;TR.IndexConstituentSectorName"",,R2C" & i * step + 1 & ")"
    
'    Application.Wait (Now + TimeValue("00:00:01"))
Next i
    
End Sub

Sub Combine()
'
' MacroDATA2 Macro
'
'Sheets("Sheet2").Select
Dim col As Integer
col = 36

Dim totalStock, startIndex As Integer
totalStock = 0
startIndex = 1



For i = 1 To 36
    Dim rs As Integer
    
    rs = Worksheets("indices_value").Cells(2, (i - 1) * 3 + 1).End(xlDown).Row
    
    If i = 1 Then
        startIndex = 1
    Else
        startIndex = totalStock + 1
    End If
    
    totalStock = totalStock + rs
    
    For j = startIndex To totalStock
        For l = 1 To 3
        
        Worksheets("indices_combine").Cells(j, l).Value = Worksheets("indices_value").Cells(j - startIndex + 1, (i - 1) * 3 + l)
        
        Next
    Next
    
    
    
    Debug.Print rs
Next


End Sub
'
'Dim i As Long
'For i = 1 To Range("K1").End(xlDown).Row
'     'your code
'Next
' '******************************************
'Dim i As Long
'For i = 1 To Range("K" & Rows.Count).End(xlUp).Row
'     'your code
'Next
' '******************************************
'Dim c As Range
'For Each c In Range("K1", Range("K1").End(xlDown))
'     'your code
'Next
' '******************************************
'Dim i As Long
'i = 1
'Do
'     'your code
'    i = i + 1
'Loop Until Range("K" & i) = ""
Sub Retrieve()

    Dim arr(984) As String

    Dim i As Integer
    i = 0
    For Each c In Worksheets("Sheet1").Range("B1001:B1984")

       arr(i) = Chr(34) & c.Value & Chr(34)

        i = i + 1
     Next c


Dim start As Integer
Dim WS As Worksheet
Set WS = Sheets.Add
WS.Select

start = 801
For i = start To 984
    Cells(1, (2 * (i - start + 1) - 1)).Select

    ActiveCell.FormulaR1C1 = _
        "=RHistory(" & arr(i - 1) & ",""TRDPRC_1.Timestamp;TRDPRC_1.Volume"",""NBROWS:20000 TIMEZONE:LON INTERVAL:5M"",,""CH:In;Fd"",R1C" & 2 * (i - start + 1) - 1 & ")"
       
Next i



End Sub

Sub Valuedata()
Dim ws1 As Worksheet
Set ws1 = ThisWorkbook.Worksheets("10")
ws1.UsedRange.Value = ws1.UsedRange.Value

End Sub

Sub CombineTables()
Dim index_table, volume_table As String
index_table = "indice"
volume_table = "volume"

'volume_sheet = Worksheets(volume_table)

Dim col As Integer
col = 36

Dim sum As Integer
col = 0

For i = 1 To 36
    Dim rs As Integer
    Dim mkt_idx As String
    mkt_idx = Worksheets(index_table).Cells(1, (i - 1) * 3 + 1)
    rs = Worksheets(index_table).Cells(2, (i - 1) * 3 + 1).End(xlDown).Row - 1
    
    sum = sum + rs
    
    
    If i = 1 Then
        startIndex = 1
    Else
        startIndex = totalStock + 1
    End If
    
    totalStock = totalStock + rs
    
    For j = startIndex To totalStock
        Worksheets("matrix").Cells(j, 1).Value = mkt_idx
        For l = 2 To 4
        
        Worksheets("matrix").Cells(j, l).Value = Worksheets(index_table).Cells(j - startIndex + 2, (i - 1) * 3 + l - 1)
        
        Next
    Next
Next
End Sub

Sub uniqueIndex()

End Sub

Sub FormMatrix()
Dim firstRow, timeStart, timesBins, stockNum, curStockNum, curWriteRow, indexRow As Integer
firstRow = 3
timeStart = 5
timeBins = 108
stockNum = 1984
curStockNum = 0
curWriteRow = 1
indexRow = 802


Dim matrix, curr, index As Worksheet
Set matrix = Worksheets("matrix5")
Set index = Worksheets("index")


'get timebins value
ReDim arrayTime(timeBins) As Date
For i = timeStart + 1 To timeStart + timeBins
    arrayTime(i - (timeStart + 1)) = matrix.Cells(1, i).Value
Next

'temp variable for read into matrix
Dim arraydate() As String
Dim curDate() As String
Dim curDay As String
Dim readDay, readTime As String
Dim dateValue As Double
Dim rowns, clos As Integer
Dim cursheet As Integer
cursheet = 5

For Sheet = cursheet To cursheet
    Set curr = Worksheets(CStr(Sheet))
    
    
    'clos = curr.Cells(1, 1).End(xlToRight).Column / 2
    clos = 200
    curDate() = Split(Now)
    curDay = curDate(0)
    Debug.Print "cloumns "; clos
    
    For l = 1 To clos
        Debug.Print l
        'Debug.Print curr.Cells(400, 400).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        curStockNum = curStockNum + 1
        Debug.Print curr.Cells(1, l * 2).Value
        
        If Not IsNumeric(curr.Cells(firstRow, l * 2 - 1)) Then
            rowns = 3
        Else
            rowns = curr.Cells(1, l * 2).End(xlDown).Row
        End If
        'start date
        
        'curDate() = Split(CDate(curr.Cells(firstRow, l * 2 - 1)))
        'Debug.Print curDate(0)
        'curDay = curDate(0)
        
        'get all date for each stock
        Dim test As Integer
        test = rowns
        For r = 3 To test
            If Not IsNumeric(curr.Cells(r, l * 2 - 1)) Then Exit For
        
            dateValue = curr.Cells(r, l * 2 - 1).Value
            
            arraydate() = Split(CDate(dateValue))
            
            readDay = arraydate(0)
            readTime = arraydate(1)
            'Debug.Print readDay
            'Debug.Print curDay
            If readDay = curDay Then
                'fill in volume
                'Debug.Print readDay
                'Debug.Print arrayTime(timeBins - 1)
                If TimeValue(readTime) > TimeValue(arrayTime(timeBins - 1)) Then
                    r = r + 1
                ElseIf TimeValue(readTime) < TimeValue(arrayTime(0)) Then
                    r = r + 1
                Else

                    For t = 108 To 1 Step -1
                         If TimeValue(arrayTime(t - 1)) = TimeValue(readTime) Then
                            'Debug.Print curr.Cells(r, l * 2)
                            matrix.Cells(curWriteRow, t + timeStart) = curr.Cells(r, l * 2)
                            'r = r + 1
                            'read new line
                            If r + 1 <= test Then
                                
                                dateValue = curr.Cells(r + 1, l * 2 - 1).Value
                                arraydate() = Split(CDate(dateValue))
                                
                                readDay = arraydate(0)
                                readTime = arraydate(1)
                                If readDay = curDay Then
                                    r = r + 1
                                    
                                Else
                                    If t <> 1 Then
                                    t = t - 1
                                        For tempt = t To 1 Step -1
                                            matrix.Cells(curWriteRow, tempt + timeStart).Value = 0
                                            'matrix.Range(Cells(curWriteRow, 1 + timeStart), Cells(curWriteRow, t - 1 + timeStart)).Value = 0
                                        Next
                                        r = r + 1
                                        t = 1
                                    End If
                                End If
                            Else
                                'If TimeValue(readTime) > TimeValue(arrayTime(0)) Then
                                If t <> 1 Then
                                    t = t - 1
                                    For tempt = t To 1 Step -1
                                        matrix.Cells(curWriteRow, tempt + timeStart).Value = 0
                                            'matrix.Range(Cells(curWriteRow, 1 + timeStart), Cells(curWriteRow, t - 1 + timeStart)).Value = 0
                                    Next
                                    'matrix.Range(Cells(curWriteRow, 1 + timeStart), Cells(curWriteRow, t - 1 + timeStart)).Value = 0
                                End If
                                t = 1
                                
                            End If
                         Else
                            matrix.Cells(curWriteRow, t + timeStart) = 0
                         End If
                         
                         If t = 1 Then
                            matrix.Cells(curWriteRow, timeStart).Value = readDay
                         End If
                         
                    Next
                    
                End If

                
                'If TimeValue(arraydate(1)) > arrayTime()
            Else
            curWriteRow = curWriteRow + 1
            
                curDate() = Split(CDate(curr.Cells(r, l * 2 - 1)))
                curDay = curDate(0)
                For c = 1 To 4
                    matrix.Cells(curWriteRow, c) = index.Cells(indexRow, c)
                
                Next
                r = r - 1
                'Debug.Print Split((CDate(arr(r))," ")
            End If
        Next
        indexRow = indexRow + 1
        Debug.Print "current write"; curWriteRow, indexRow
    Next

Next

'ActiveWorkbook.Save

End Sub



