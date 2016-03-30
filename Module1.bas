Attribute VB_Name = "Module1"
Sub GetDATA()
'add a table of courses into an Excel workbook
'the website containing the files listing courses
Dim ROW As Long, Arow As Long, symbol As String, URL As String
Dim str As String


With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).ROW
End With
Arow = LastRow
'we decide how many company data that we need.if we want all the data ,we should make loop condition like "for j=2 to Arow",
'while if we want limited number of company, the loop condition like "for j=2 to number"
For j = 2 To 15
Range("F:G").Clear
ROW = 1
str = Range("A" & j).Value
symbol = regex(str, "\d{4}", "$0")
For i = 1 To 2
'Const prefix As String = "http://finance.yahoo.com/q?uhb=uh3_finance_vert&fr=&type=2button&s=" &symbol&".HK%2C"
URL = "http://finance.yahoo.com/q?uhb=uh3_finance_vert&fr=&type=2button&s=" & symbol & ".HK%2C"
'Const FileName As String = "microsoft-excel-advanced"
Dim qt As QueryTable
Dim ws As Worksheet
'using a worksheet variable means autocompletion works better
Set ws = ActiveSheet
'set up a table import (the URL; tells Excel that this query comes from a website)
Set qt = ws.QueryTables.Add( _
Connection:="URL;" & URL, _
Destination:=Range("F" & ROW))
'tell Excel to refresh the query whenever you open the file
qt.RefreshOnFileOpen = False
'giving the query a name can help you refer to it later
qt.Name = "ExcelAdvancedCoursesFromWiseOwl"
'you want to import column headers
qt.FieldNames = True
'need to know name or number of table to bring in
'(we'll bring in the first table)
qt.WebSelectionType = xlSpecifiedTables
qt.WebTables = i + 1
'import the data
qt.Refresh BackgroundQuery:=False
With ActiveSheet
        LastRow = .Cells(.Rows.Count, "F").End(xlUp).ROW
End With
ROW = LastRow + 2
Next

Range("B" & j) = Range("G1").Value
Range("C" & j) = Range("G2").Value
Range("D" & j) = Range("G9").Value
Range("E" & j) = Range("G11").Value
Next


End Sub

Function regex(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then
        regex = False
    Else
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.Pattern = "\$" & replaceNumber

            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                    regex = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regex = outputPattern
    End If
End Function

