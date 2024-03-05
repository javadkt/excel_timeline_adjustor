Sub UpdateDates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim effortCol As Long
    Dim startDateCol As Long
    Dim endDateCol As Long
    Dim row As Long
    Dim tempEndDate As Date
    Dim sundayCount As Integer

    Dim beforeDecimal As Double
    Dim afterDecimal As Double
    Dim hasDecimal As Boolean

    ' Set the worksheet And column indices
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" To your sheet name
    effortCol = 3 ' Column C For Effort
    startDateCol = 4 ' Column D For Start Date
    endDateCol = 5 ' Column E For End Date

    ' Find the last row With data in column C (Effort column)
    lastRow = ws.Cells(ws.Rows.Count, effortCol).End(xlUp).row

    ' Loop through rows (skip first row As it contains headers)
    For row = 2 To lastRow
        Dim effortVal As Double
        effortVal = ws.Cells(row, effortCol).Value
        If Not IsEmpty(effortVal) And IsNumeric(effortVal) Then
            If row > 2 Then


                Dim decimalSumValue As Double
                Dim currentRow As Integer
                Dim endDate As Variant

                ' Get the end date from row - 1
                endDate = ws.Cells(row - 1, endDateCol).Value

                ' Start from row - 1 And move upwards Until reaching a row With a different end date
                currentRow = row - 1
                decimalSumValue = 0

                Do While ws.Cells(currentRow, endDateCol).Value = endDate And currentRow > 1
                    ' Add the value in effortCol To the sum
                    decimalSumValue = decimalSumValue + CDbl(ws.Cells(currentRow, effortCol).Value)
                    ' Move To the previous row
                    currentRow = currentRow - 1
                Loop

                ' Check If the sum is a whole number
                If decimalSumValue = Int(decimalSumValue) Then
                    ws.Cells(row, startDateCol).Value = Format(DateAdd("d", 1, ws.Cells(row - 1, effortCol).Value), "yyyy-mm-dd")
                Else
                    ws.Cells(row, startDateCol).Value = ws.Cells(row - 1, endDateCol).Value
                End If

 
            End If



            'Extracting Decimals Starts'

            beforeDecimal = Int(effortVal)
            afterDecimal = effortVal - beforeDecimal

            If afterDecimal > 0 Then
                hasDecimal = True
            Else
                hasDecimal = False
            End If

            ' ' Output the results
            ' MsgBox "Before Decimal: " & beforeDecimal & vbCrLf & _
            ' "After Decimal: " & afterDecimal & vbCrLf & _
            ' "Has Decimal: " & hasDecimal
            ' 'Extracting Decimals Ends'


            If hasDecimal Then
                ' End Date is Start Date + Effort - 1

                If beforeDecimal = 0 Then
                    tempEndDate = ws.Cells(row, startDateCol).Value
                    'more To Do'
                Else
                    'Adding For before decimal'
                    tempEndDate = DateAdd("d", beforeDecimal - 1, ws.Cells(row, startDateCol).Value)

                    'Adding For after decimal'
                    If afterDecimal <> 0 Then
                        tempEndDate = DateAdd("d", 1, ws.Cells(row, startDateCol).Value)
                    End If
                End If

            Else
                tempEndDate = DateAdd("d", ws.Cells(row, effortCol).Value - 1, ws.Cells(row, startDateCol).Value)
            End If


            sundayCount = 0 ' Reset Sunday count

            ' Loop through each day in the date range
            For i = ws.Cells(row, startDateCol).Value To tempEndDate
                If Weekday(i) = 1 Then ' Check If it's Sunday (Weekday Function returns 1 For Sunday)
                    sundayCount = sundayCount + 1 ' Increment Sunday count
                End If
            Next i

            ' Adjust End Date For multiple Sundays
            tempEndDate = DateAdd("d", sundayCount, tempEndDate)

            ws.Cells(row, endDateCol).Value = Format(tempEndDate, "yyyy-mm-dd")
        End If
    Next row

    ' Refresh the screen To reflect changes
    Application.ScreenUpdating = True

    MsgBox "Dates updated successfully."
End Sub








