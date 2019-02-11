

Sub Surfy()

Rows(1).EntireRow.Delete
Columns(1).EntireColumn.Delete
Columns(1).EntireColumn.Delete

Dim r As Integer
For r = Sheet1.UsedRange.Rows.Count To 1 Step -1
    If (Cells(r, "B") = "Best Buy Canada" Or Cells(r, "B") = "Canada Computers" Or Cells(r, "B") = "Aniks Appliances Inc" Or Cells(r, "B") = "Lowes.ca" Or Cells(r, "B") = "HomeDepot.ca" Or Cells(r, "B") = "Bloor Dovercourt Appliances" Or Cells(r, "B") = "Goemans Appliances" Or Cells(r, "B") = "Universal Appliances" Or Cells(r, "B") = "TheBay.com" Or Cells(r, "B") = "Coast Appliances" Or Cells(r, "B") = "Total Appliance Centre" Or Cells(r, "B") = "RONA" Or Cells(r, "B") = "Corbeil Appliances - Ontario" Or Cells(r, "B") = "TeleTime Appliances" Or Cells(r, "B") = "Crawford Appliance Centre") Then
        Sheet1.Rows(r).EntireRow.Delete
    End If
Next

'Columns(2).EntireColumn.Delete

For Each c In Range("E1:E200")
    If InStr(c.Value, " ") > 0 Then
        c.Value = Right(c.Value, InStr(c.Value, " "))
    End If
Next c

Dim myRange As Range
Dim myCell As Range
Set myRange = Range("D2:D200")
For Each myCell In myRange
    If Not myCell Like "*Canadian-Appliance-Source*" Then
        myCell.Value = ""
    End If
Next myCell

Dim rng As Range
    Dim i As Long
    Set rng = ThisWorkbook.ActiveSheet.Range("D1:D200")
    Set rng2 = ThisWorkbook.ActiveSheet.Range("C1:C200")
    Set rng3 = ThisWorkbook.ActiveSheet.Range("E1:E200")
    With rng
        For i = .Rows.Count To 1 Step -1
            If .Item(i) = "" And rng2.Item(i) = "" Then
                .Item(i).EntireRow.Delete
            End If
        Next i
  
        For i = rng2.Rows.Count To 1 Step -1
            If rng2.Item(i) <> "" Then
                rng3.Item(i).Value = rng2.Item(i)
            End If
        Next i
    End With

Columns(2).EntireColumn.Delete
Columns(3).EntireColumn.Delete
Columns(2).EntireColumn.Delete
Columns(2).EntireColumn.Insert

End Sub


