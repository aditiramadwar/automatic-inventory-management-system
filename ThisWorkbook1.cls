VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Workbook_Open()

run = False
'Sheet2.Activate

cell_row = Selection.Row                                    'see which row is selected
cell_col = Selection.Column
    
Dim btn As button
Set btn = ActiveSheet.Buttons.Add(Cells(cell_row, cell_col + 2).Left, Cells(cell_row, cell_col + 2).Top, Range("C1").Width, Range("C1").Height)
       
If btn.Name = "Button 1" Then
    Debug.Print btn.Name
     
'creates all the settings for the button
    With btn
        .Name = "clickit"
        .OnAction = "start_stop"
        .Caption = "START"
        .Font.Bold = True
        
    End With
Else

    btn.Delete
    
End If

'initial cells
Range("A1").Value = "Code"
Range("B1").Value = "Barcode"

End Sub


