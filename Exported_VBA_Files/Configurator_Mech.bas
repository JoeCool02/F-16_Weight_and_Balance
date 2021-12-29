Attribute VB_Name = "Configurator_Mech"
Dim oxlWorksheet As Excel.Worksheet         'Calculations Worksheet
Dim oxlWorksheetConfig As Excel.Worksheet   'Configuration Worksheet

'Ranges for Update_Colored_Stores Sub
Dim oxlRange As Excel.Range, _
    oxlRangeConfig As Excel.Range, _
    oxlRangeManShapes As Excel.Range

Dim Counter As Integer                      'Counter for Update_Colored_Stores Sub
   
'Shape variables
Dim shpStation As Shape
Dim Shape_Array(25) As String
    
Sub On_Stores_Dropdown_Click()

    Update_Colored_Stores
    
    Arrange_Stores
    Find_Worst_CG
    Copy_Worst_Case_Jett
    ID_Mission
    Fill_Remarks
    
End Sub

Sub On_AME_Dropdown_Click()
    
    Arrange_Stores
    Find_Worst_CG
    Copy_Worst_Case_Jett
    ID_Mission
    Fill_Remarks
    
    Fill_Stores_From_AME
    Update_Stores_Dropdowns
    
    Update_Colored_Stores
    
End Sub

Sub Update_Colored_Stores()

    Dim Manual_Shapes As New Collection

    Set oxlWorksheet = Worksheets("Calculations")
    Set oxlWorksheetConfig = Worksheets("Configurator")
    Set oxlRange = oxlWorksheet.Range("AC3:AC28")
    Set oxlRangeShp = oxlWorksheet.Range("CF3:CF28")
    Set oxlRangeConfig = oxlWorksheetConfig.Range("A52:A63")
    Set oxlRangeManShapes = oxlWorksheet.Range("DH3:DH11")
    
    For Each C In oxlRangeManShapes.Cells
        
        Manual_Shapes.Add C.Offset(0, 1).Value, C.Text
        Set_Color_InActive (C.Offset(0, 1).Value)
        
    Next C
    
    Counter = 0
    
    For Each var In oxlRangeShp.Value2
        Shape_Array(Counter) = var
        Counter = Counter + 1
    Next var
    
    Counter = 0

    For Each C In oxlRange.Cells
                
        If C.Value = "0" Then
            Set_Color_InActive (Shape_Array(Counter))
        Else
            Set_Color_Active (Shape_Array(Counter))
        End If
        
        Counter = Counter + 1

    Next C
    
    For Each C In oxlRangeConfig.Cells
    
        If C.Value <> "" Then
        
            Set_Color_Active (Manual_Shapes(C.Offset(0, 4)))
            
        End If
        
    Next C

End Sub

Sub Set_Color_Active(Shape_Name As String)

    Set oxlWorksheet = Worksheets("Configurator")
    Set shpStation = oxlWorksheet.Shapes(Shape_Name)

    With shpStation
        .Visible = msoTrue
    End With

End Sub

Sub Set_Color_InActive(Shape_Name As String)

    Set oxlWorksheet = Worksheets("Configurator")
    Set shpStation = oxlWorksheet.Shapes(Shape_Name)

    With shpStation
        .Visible = msoFalse
    End With

End Sub

Sub Update_Ret_Jett_Exp()

    '********************DEPRECATED********************************************
    '*************OBSOLETE AS OF VERSION 6.0***********************************

    Dim oxlWorksheet As Excel.Worksheet, oxlWorksheetConfig As Excel.Worksheet
    Dim oxlRange As Excel.Range
    Dim Shape_Array(19) As String
    Dim Counter As Integer
    Dim Status_Textbox As Shape
    
    Set oxlWorksheet = Worksheets("Calculations")
    Set oxlWorksheetConfig = Worksheets("Configurator")
    Set oxlRange = oxlWorksheet.Range("AG3:AG22")
    Set oxlRangeShp = oxlWorksheet.Range("CX3:CX22")
    
    Counter = 0
    
    For Each var In oxlRangeShp.Value2
        Shape_Array(Counter) = var
        Counter = Counter + 1
    Next var
    
    Counter = 0

    For Each C In oxlRange.Cells
                
        Set Status_Textbox = oxlWorksheetConfig.Shapes(Shape_Array(Counter))
        
        If Shape_Array(Counter) <> "" Then
            If C.Value = "1" Then
                Status_Textbox.TextEffect.Text = "Retained"
            ElseIf C.Value = "2" Then
                Status_Textbox.TextEffect.Text = "Jettisonable"
            ElseIf C.Value = "3" Then
                Status_Textbox.TextEffect.Text = "Expendable"
            End If
        End If
        
        Counter = Counter + 1

    Next C

End Sub

Sub Update_Check_Boxes()

'*************************DEPRECATED**************************
'***************NOT USED IN V6.1 AND LATER********************

    Dim oxlWorksheet As Excel.Worksheet
    Dim oxlRange1 As Excel.Range, oxlRange2 As Excel.Range
    Dim Counter As Integer

    Set oxlWorksheet = Worksheets("Calculations")
    Set oxlRange1 = oxlWorksheet.Range("AC3:AC22")
    Set oxlRange2 = oxlWorksheet.Range("AG3")
    
    Counter = 0
    
    For Each C In oxlRange1.Cells
    
        If C.Value = 0 Then
            oxlRange2.Offset(Counter, 0) = "1"
        End If
        
        Counter = Counter + 1
        
    Next C

End Sub

Sub Copy_Limit_Formats()
'**********************DEPRECATED**************************

    Dim oxlConfig As Excel.Worksheet
    Dim oxlCalcs As Excel.Worksheet
    
    Set oxlConfig = Worksheets("Configurator")
    Set oxlCalcs = Worksheets("Calculations")
    
    oxlConfig.Range("S11").Interior.Color = oxlCalcs.Range("A20").Interior.Color

End Sub
