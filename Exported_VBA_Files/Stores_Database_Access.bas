Attribute VB_Name = "Stores_Database_Access"
Dim oxlSheetCalc As Excel.Worksheet
Dim oxlSheetConfig As Excel.Worksheet
Dim oxlSheetStores As Excel.Worksheet

Dim oxlRange As Excel.Range
Dim oxlRangeCalc As Excel.Range

Dim oWorkSheet As Excel.Worksheet

Dim oDropDownRng As Excel.Range
Dim oDropDownShp As Shape

Dim Path As String
Dim sSQL As String

Dim Connection As Object
Dim Recordset As Object

Sub Quick_Stores_Update()

    Fill_AME_From_Database
    Fill_Stores_From_AME
    Update_Stores_Dropdowns

End Sub

Sub Update_Stores_Dropdowns()

    Set oxlSheetCalc = Worksheets("Calculations")
    Set oxlSheetConfig = Worksheets("Configurator")
    Set oDropDownRng = oxlSheetCalc.Range("DA3:DA28")
    
    For Each C In oDropDownRng.Cells
    
        If C.Value <> "" Then
            
            Set oDropDownShp = oxlSheetConfig.Shapes(C.Value)
            oDropDownShp.ControlFormat.ListFillRange = C.Offset(0, 2).Value
            
        End If
        
    Next C
    
End Sub

Sub Fill_AME_From_Database()

    Set oxlSheetStores = Worksheets("Stores")
    Set oxlRange = oxlSheetStores.Range("A202:BC470")
    oxlRange.Value = ""
    Set oxlRange = oxlSheetStores.Range("A502:BC600")
    oxlRange.Value = ""

    ReadStationAME "'1/9'", "A202", "Stores"
    ReadStationAME "'2/8'", "F202", "Stores"
    ReadStationAME "'3/7'", "K202", "Stores"
    ReadStationAME "'4/6'", "P202", "Stores"
    
    ReadStationAME "'1/9'", "AY202", "Stores"
    ReadStationAME "'2/8'", "AT202", "Stores"
    ReadStationAME "'3/7'", "AO202", "Stores"
    ReadStationAME "'4/6'", "AJ202", "Stores"
    
    ReadStationAME "'5'", "Z202", "Stores"

End Sub

Sub Fill_Stores_From_AME()
    
    Set oxlSheetStores = Worksheets("Stores")
    Set oxlRangeStores = oxlSheetStores.Range("A3:BC199")
    Set oxlSheetCalc = Worksheets("Calculations")
    Set oxlRangeCalc = oxlSheetCalc.Range("DD3:DD17")
    
    oxlRangeStores.Value = ""
    
    For Each C In oxlRangeCalc.Cells
    
        Debug.Print "Station = " & C.Value
        Debug.Print "AME = " & C.Offset(0, 1).Value
        Debug.Print "Output Location = " & C.Offset(0, 2).Value
        
        ReadStoresByAME C.Value, C.Offset(0, 1).Value, C.Offset(0, 2).Value, "Stores"

    Next C

End Sub

Sub ReadStationAME(AME_Station As String, Write_Cell As String, Write_Sheet As String)

    sSQL = "SELECT Stores.Store_Name, Stores.Short_Name, [Stores]![Store_Weight]*[Stores]![Quantity] AS Total_Weight, [Total_Weight]*[Stores]![FS_Arm]/100 AS Lon_MOM, [Total_Weight]*[Stores]![BLS_Arm]/100 AS Lat_MOM " & _
           "FROM Relationships INNER JOIN Stores ON Relationships.Child = Stores.ID " & _
           "WHERE (((Relationships.Parent)=0) AND ((Stores.Station)=" & AME_Station & "));"
           
    Set Connection = New ADODB.Connection
    Set Recordset = New ADODB.Recordset
    Set oWorkSheet = Worksheets(Write_Sheet)
    
    Path = oWorkSheet.Application.ActiveWorkbook.Path
    
    Connection.Provider = "Microsoft.ACE.OLEDB.12.0"
    Connection.ConnectionString = "data source=" & Path & "\StoresDatabase.accdb;Persist Security Info=False"
    Connection.Mode = adModeReadWrite
    Connection.Open
            
    Recordset.Open sSQL, Connection
    oWorkSheet.Range(Write_Cell).CopyFromRecordset Recordset
    
    Recordset.Close
    Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing
    
    Exit Sub
    
End Sub

Sub ReadStoresByAME(AME_Station As String, AME_Name As String, Write_Cell As String, Write_Sheet As String)

    sSQL = "SELECT Stores_1.Store_Name, Stores_1.Short_Name, [Stores_1]![Quantity]*[Stores_1]![Store_Weight] AS Total_Weight, [Total_Weight]*[Stores_1]![FS_Arm]/100 AS Lon_MOM, [Total_Weight]*[Stores_1]![BLS_Arm]/100 AS Lat_MOM " & _
            "FROM (Stores INNER JOIN Relationships ON Stores.ID = Relationships.Parent) INNER JOIN Stores AS Stores_1 ON Relationships.Child = Stores_1.ID " & _
            "WHERE (((Stores.Store_Name)= '" & AME_Name & "') AND ((Stores.Station)= '" & AME_Station & "'));"

    Set Connection = New ADODB.Connection
    Set Recordset = New ADODB.Recordset
    Set oWorkSheet = Worksheets(Write_Sheet)
    
    Path = oWorkSheet.Application.ActiveWorkbook.Path
    
    Connection.Provider = "Microsoft.ACE.OLEDB.12.0"
    Connection.ConnectionString = "data source=" & Path & "\StoresDatabase.accdb;Persist Security Info=False"
    Connection.Mode = adModeReadWrite
    Connection.Open
    
    Recordset.Open sSQL, Connection
    
    oWorkSheet.Range(Write_Cell).CopyFromRecordset Recordset
    
    Recordset.Close
    Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing

End Sub
