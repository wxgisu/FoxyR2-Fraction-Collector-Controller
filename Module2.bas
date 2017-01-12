Attribute VB_Name = "Module2"

Sub RunMeFirst()
' This sub routine is used to setup the control sheet for running foxy R2 Fraction collector
' Creates a new sheet with controls to run Foxy
' This sheet is used to input fraction timing.
    ' This was created using record macro feature and not the best code, but sets up the page pretty quick. 
If Not SheetExists("FoxyCol") Then
    
    ThisWorkbook.Sheets.Add.Activate
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "End Time"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Next Run"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "State"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Next Call 0=StartFrac 1=MoveFrac"
    Columns("B:B").Select
    Rows("5:5").RowHeight = 25.5
    Columns("A:A").ColumnWidth = 10.57
    Columns("A:A").ColumnWidth = 11.29
    Range("A5").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").ColumnWidth = 15
    Rows("5:5").RowHeight = 41.25

    
    Range("B12").Select
    ActiveSheet.Buttons.Add(84, 100.5, 45.75, 14.25).Select
    Selection.OnAction = _
        "CommandButton1_Click"
    Selection.Characters.Text = "StartFrac" & Chr(10) & ""
    With Selection.Characters(Start:=1, length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
       
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "Input Values"
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "Foxy IP Address"
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "Total Time (hr:mm:ss)"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Sample Interval "
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Sampling Time (hr:mm:ss)"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Sample Interval (hr:mm:ss) "
    Range("B16").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Next Tube No"
   
    
    Range("B16").Select
    ActiveCell.FormulaR1C1 = "1"
     Range("B12:B15").Select
    Selection.NumberFormat = "@"
    Range("B12:B16").Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "129.186.18.31"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "Make sure the computer and Foxy are on the same subnet and first 3 numbers in the Ip address match."
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "01:00:00"
    Range("B14").Select
    ActiveCell.FormulaR1C1 = "00:05:00"
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "00:00:30"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = "Subtracts Sampling time"
    Range("B21").Select
    Columns("B:B").ColumnWidth = 14.57
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Macro Calculated Values for reference"
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "Total Time "
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Waiting time interval"
    Range("A22").Select
    ActiveCell.FormulaR1C1 = "Sampling time"
    Range("A25").Select
    ActiveCell.FormulaR1C1 = "Start_frac_Counter"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "Mov_Frac_Counter"
    Range("A19:B21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("C14").Select
    Selection.Copy
    Range("C20").Select
    ActiveSheet.Paste
 
    ActiveSheet.Buttons.Add(364.5, 15, 90, 15).Select
 
    Selection.OnAction = _
        "outletCleanSetup"
    Selection.Characters.Text = "Clean Outlet"
    With Selection.Characters(Start:=1, length:=12).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("H4").Select
    ActiveSheet.Buttons.Add(364.5, 45, 90, 15).Select
    Selection.OnAction = _
        "Stop_Cleanout"
    Selection.Characters.Text = "Stop Clean"
    With Selection.Characters(Start:=1, length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("G4").Select
    Range("I5").Select
    ActiveSheet.Buttons.Add(411, 100.5, 48.75, 15).Select
    Selection.OnAction = _
        "StopFrac_UserClick"
    Selection.Characters.Text = "Stop Frac"
    With Selection.Characters(Start:=1, length:=9).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
 
    Cells.Select

Application.ActiveSheet.Name = "FoxyCol"
 Range("B20:B22").Select
 Selection.NumberFormat = "h:mm:ss"
Application.Sheets("FoxyCol").Protect , userinterfaceonly:=True
Else
    MsgBox ("Sheet with name Foxycol already exist, remove or remain it before running this")
End If
    

End Sub

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
