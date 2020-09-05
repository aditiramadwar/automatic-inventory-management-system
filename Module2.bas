Attribute VB_Name = "Module2"
Public nextScheduledTime As Date
Public run As Boolean

Sub start_stop()
If ActiveSheet.Buttons("clickit").Caption = "START" Then
        run = True
        ActiveSheet.Buttons("clickit").Caption = "STOP"
'        ActiveSheet.Buttons("clickit").BackColor = RGB(150, 0, 0)
    Else
        run = False
        ActiveSheet.Buttons("clickit").Caption = "START"
 '       ActiveSheet.Buttons("clickit").BackColor = RGB(0, 150, 0)
    End If
    Debug.Print run
    
Call PrintBarcode
End Sub

Sub PrintBarcode()
'run = False
    
    If run = True Then
        cell_row = Selection.Row                                    'see which row is selected
        cell_col = Selection.Column
        
        If cell_row = 1 Then
            Cells(cell_row + 1, cell_col).Select
            cell_row = Selection.Row
        End If
            
        b = Len(Cells(cell_row - 1, cell_col))                      'check the length of the text in the cell
        c = Cells(1, 15).Value
        
        'If b = 19 Then                                             'check if proper barcode has been scanned similar to:EM0201917090004C7_R
        If b > 1 Then
            If c = cell_row - 1 Then                                'check if selected cell is different
                Debug.Print Cells(1, 15)                             'print row number
                Debug.Print Cells(cell_row - 1, cell_col + 1)       'print barcode data
                Call cell_format(cell_row, cell_col)                'call sub for cell format settings
                Call PrintSet(cell_row, cell_col)                   'call sub for printing settings
                'Cells(cell_row - 1, cell_col + 1).PrintOut         'print the barcode from the cell next to it'
            End If
            
        End If
        Cells(1, 15).Value = cell_row                                   'print current row number in cell
        CreateNewSchedule
    End If
    
    ActiveSheet.Buttons("clickit").Top = Rows(ActiveCell.Row - 1).Top
    ActiveSheet.Buttons("clickit").Left = Columns(ActiveCell.Column + 2).Left
    ActiveSheet.Buttons("clickit").Width = Columns(ActiveCell.Row).Width
    ActiveSheet.Buttons("clickit").Height = Rows(ActiveCell.Column + 2).Height
    
End Sub

Private Sub CreateNewSchedule()
 nextScheduledTime = DateAdd("s", 1, Now)
    Application.OnTime EarliestTime:=nextScheduledTime, Procedure:="PrintBarcode", Schedule:=True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
On Error Resume Next
    Application.OnTime EarliestTime:=nextScheduledTime, Procedure:="PrintBarcode", Schedule:=False
End Sub

Sub cell_format(cell_row, cell_col)
'set up font and size
    Cells(cell_row - 1, cell_col).Font.Name = "Calibri"                                     'code font
    'Cells(cell_row - 1, cell_col + 1).Font.Size = 10                                       'barcode font size for M Barcode font
    Cells(cell_row - 1, cell_col + 1).Font.Size = 8                                         'barcode font size for XL and XXL
    'change font name as per label and coverage requirement
    'Cells(cell_row - 1, cell_col + 1).Font.Name = "IDAHC39M Code 39 Barcode"               'barcode font medium
    Cells(cell_row - 1, cell_col + 1).Font.Name = "IDAutomationSYHC39XL Demo Sym"           'barcode font XL
    'Cells(cell_row - 1, cell_col + 1).Font.Name = "IDAutomationSYHC39XXL Demo Sym"         'barcode font XXL
    Cells(cell_row - 1, cell_col + 1).FormulaR1C1 = "=""(""&RC[-1]&"")"""                   'set formula for barcode
    Columns(cell_col).AutoFit
    'cell formatting settings
    With Cells(cell_row - 1, cell_col + 1)
        .Columns.AutoFit
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlDistributed
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
End Sub
Sub PrintSet(cell_row, cell_col)


'printing settings
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        '.LeftMargin = Application.InchesToPoints(0.3)
        .LeftMargin = Application.InchesToPoints(0.118110236220472)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0.3)
        '.TopMargin = Application.InchesToPoints(0.118110236220472)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 203
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = 260
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = True
        .Zoom = 69
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
End Sub


