Attribute VB_Name = "Tail_Number"
Sub DropDown69_Change()

    Dim oxlWorksheet As Excel.Worksheet
    Set oxlWorksheet = Worksheets("Calculations")
    
    With oxlWorksheet
        .Range("D3").Value = .Range("B6").Value
    End With
    
    On_Stores_Dropdown_Click
    
End Sub
