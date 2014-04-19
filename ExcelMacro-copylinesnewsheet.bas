Attribute VB_Name = "Module1"


Sub copylinesnewsheet()

'works in Windows Excel 2010
'Ctrl+h shortcut
'will copy first two lines of each sheet and paste in subsequent lines on Sheet1

    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets

        ws.Activate

        Range("A1:A2").Select

        Selection.Copy

        Sheets("Sheet1").Select

        lMaxRows = Cells(Rows.Count, "B").End(xlUp).Row

        Range("B" & lMaxRows + 1).Select

        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _

        False, Transpose:=False

        Range("A1:A2").Select

    Next

End Sub

