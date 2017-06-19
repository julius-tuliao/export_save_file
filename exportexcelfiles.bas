Attribute VB_Name = "export"
Option Explicit

Sub Export_Click()
Dim fpath, fname1 As String
Dim newbook As Workbook
Dim lastrow As Long
Application.DisplayAlerts = False
With ThisWorkbook.Worksheets("Data")
    lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With


fpath = "F:\Current Project\Collection"
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & ".csv"




ThisWorkbook.Worksheets("Data").Range("A1:Q7000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close


ThisWorkbook.Worksheets("Data").Range("A7001:Q14000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & "2 .csv"
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close


ThisWorkbook.Worksheets("Data").Range("A14001:Q21000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & "3 .csv"
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close


ThisWorkbook.Worksheets("Data").Range("A21001:Q28000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & "4 .csv"
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close


ThisWorkbook.Worksheets("Data").Range("A28001:Q35000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & "5 .csv"
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close

ThisWorkbook.Worksheets("Data").Range("A35001:Q42000").Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
fname1 = "PSB DATA 1 " & Format(Now(), "mmddyy") & "6 .csv"
newbook.SaveAs Filename:=fpath & "\" & fname1, FileFormat:=xlCSV, CreateBackup:=False
newbook.Close



Dim ws As Worksheet
 Dim wa As Worksheet
For Each ws In Worksheets
If ws.Name <> "Main" And ws.Name <> "Data" Then ws.Delete
Next

    With ThisWorkbook
        Set wa = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        wa.Name = "Raw Data"
    End With

Application.DisplayAlerts = True
End Sub
