Attribute VB_Name = "Configurator_Quick_Buttons"
Sub Clear_Config()

    Dim oxlWorksheet As Excel.Worksheet, oxlWorksheetConfig As Excel.Worksheet
    Dim oxlRange As Excel.Range, oxlChaffFlare As Excel.Range, oxlBackseater As Excel.Range, oxlManualTailNumber As Range
    Dim oxlManualStores As Excel.Range, oxlManualStoresStation As Excel.Range, oxlManualStoresJett As Excel.Range
    Dim oxlForceSA As Excel.Range
    
    Set oxlWorksheet = Worksheets("Calculations")
    Set oxlWorksheetConfig = Worksheets("Configurator")
    Set oxlRange = oxlWorksheet.Range("AB3:AB28")
    Set oxlChaffFlare = oxlWorksheet.Range("AA62")
    Set oxlBackseater = oxlWorksheet.Range("AT11")
    Set oxlManualStores = oxlWorksheetConfig.Range("A52:D63")
    Set oxlManualStoresStation = oxlWorksheetConfig.Range("E52:E63")
    Set oxlManualStoresJett = oxlWorksheetConfig.Range("F52:F63")
    Set oxlManualTailNumber = oxlWorksheetConfig.Range("A66")
    Set oxlForceSA = oxlWorksheet.Range("BY5")
    
    For Each C In oxlRange.Cells
        C.Value = 1
        C.Offset(0, 5).Value = 1
    Next C
    
    oxlChaffFlare.Value = "FALSE"
    oxlBackseater.Value = "FALSE"
    oxlForceSA.Value = "FALSE"
    
    oxlManualStores.Value = ""
    oxlManualStoresStation.Value = 1
    oxlManualStoresJett.Value = 1
    
    oxlManualTailNumber.Value = ""
    
    On_Stores_Dropdown_Click
    Quick_Stores_Update

End Sub

Sub Quick_1_Bag()

    'One button add centerline tank
    Dim oxlWorksheet As Excel.Worksheet
    Set oxlWorksheet = Worksheets("Calculations")
    
    'Add pylon
    oxlWorksheet.Range("AB16").Value = 4
    On_AME_Dropdown_Click
    
    'Add tank and set jettisonable
    oxlWorksheet.Range("AB15").Value = 4
    oxlWorksheet.Range("AG15") = 2
    On_Stores_Dropdown_Click

End Sub

Sub Quick_2_Bag()

    'One button add two wing tanks
    Dim oxlWorksheet As Excel.Worksheet
    Set oxlWorksheet = Worksheets("Calculations")
    
    'Add pivot balls
    oxlWorksheet.Range("AB19").Value = 13
    oxlWorksheet.Range("AB12").Value = 13
    
    'Add tank and set jettisonable
    oxlWorksheet.Range("AB18").Value = 2
    oxlWorksheet.Range("AG18") = 2
    oxlWorksheet.Range("AB11").Value = 2
    oxlWorksheet.Range("AG11") = 2
    
    On_AME_Dropdown_Click
    On_Stores_Dropdown_Click

End Sub

Sub Copy_Sta1_to_Sta9()

    Dim oxlWorksheet As Excel.Worksheet
    Set oxlWorksheet = Worksheets("Calculations")
    'Add set stores symmetric
    oxlWorksheet.Range("AB27:AB28").Value2 = oxlWorksheet.Range("AB3:AB4").Value2
        
    On_AME_Dropdown_Click
    On_Stores_Dropdown_Click
    
End Sub

Sub Copy_Sta2_to_Sta8()

    Dim oxlWorksheet As Excel.Worksheet
    Set oxlWorksheet = Worksheets("Calculations")
    'Add set stores symmetric
    oxlWorksheet.Range("AB24:AB26").Value2 = oxlWorksheet.Range("AB5:AB7").Value2
        
    On_AME_Dropdown_Click
    On_Stores_Dropdown_Click

End Sub
