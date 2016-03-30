Attribute VB_Name = "Module2"
Sub GetcompanyList1()
Range("A:H").Clear
Range("B1").Value = "Prev Close"
Range("C1").Value = "Open"
Range("D1").Value = "Day's Range"
Range("E1").Value = "Volume"

'add a table of courses into an Excel workbook
'the website containing the files listing courses
Const prefix As String = "https://en.wikipedia.org/wiki/List_of_companies_listed_on_the_Hong_Kong_Stock_Exchange"
'Const FileName As String = "microsoft-access-vba"
Dim qt As QueryTable
Dim ws As Worksheet
'using a worksheet variable means autocompletion works better
Set ws = ActiveSheet
Dim LastRow As Long
Dim ROW As Long
ROW = 2
'we define the number of seriers of symbol, we change the loop condition to get data
For i = 3 To 4
Set qt = ws.QueryTables.Add( _
Connection:="URL;" & prefix, Destination:=Cells(ROW, 1))
'set up a table import (the URL; tells Excel that this query comes from a website) ,Destination:=Range("A2")
'tell Excel to refresh the query whenever you open the file
'qt.RefreshOnFileOpen = True
'giving the query a name can help you refer to it later
'qt.Name = "ExcelAdvancedCoursesFromWiseOwl"
'you want to import column headers
qt.FieldNames = True
'need to know name or number of table to bring in
'(we'll bring in the first table)
qt.WebSelectionType = xlSpecifiedTables


'Range("B2").Value = qt.ResultRange


qt.WebTables = i

'qt.Destination = Cells(2, 2)
qt.Refresh BackgroundQuery:=False

With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).ROW
End With
ROW = LastRow + 1

Next

'qt.WebTables = 3
'import the data
'qt.Refresh BackgroundQuery:=False


End Sub



