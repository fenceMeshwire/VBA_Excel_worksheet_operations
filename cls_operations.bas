Option Explicit

' ------------------------------------------------------
Public Function create_sheets()

Dim intCounterWorksheets As Integer
Dim lngRow, lngRowMax As Long
Dim strTicker As String

intCounterWorksheets = ThisWorkbook.Worksheets.Count
lngRowMax = master.UsedRange.Rows.Count

For lngRow = 2 To lngRowMax
  strTicker = master.Cells(lngRow, 4).Value
  If Not table_exists(strTicker) Then
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = strTicker
  End If
Next lngRow

End Function

' ------------------------------------------------------
Function table_exists(strTicker As String) As Boolean

Dim wksSheet As Worksheet

For Each wksSheet In ThisWorkbook.Worksheets
  If wksSheet.Name = strTicker Then
    table_exists = True
    Exit For
  End If
Next wksSheet

End Function
