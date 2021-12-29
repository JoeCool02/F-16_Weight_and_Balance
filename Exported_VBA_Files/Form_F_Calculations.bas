Attribute VB_Name = "Form_F_Calculations"
'Fill Remarks Variables
Dim oFormF As Excel.Worksheet
Dim oCalcs As Excel.Worksheet

Dim oFormFRange As Excel.Range
    
Dim Counter As Integer
Dim Pod_Config As String
Dim Landing_Cat As String
Dim All_NWS As String

'Arrange Stores Variables
Dim oWorkSht As Excel.Worksheet, _
    oWorkShtConfig As Excel.Worksheet, _
    oWorkShtStores As Excel.Worksheet
    
Dim Exp_Lead_Text As String
Dim Retained_Row_Count As Integer, _
    Jett_Row_Count As Integer
Dim Retained_Column As Integer, _
    Jett_Column As Integer
    
'Find Worst CG Variables
Dim Array_Length As Integer, _
    Column As Integer, _
    Total_Weight As Integer, _
    Total_Lon_Moment As Integer, _
    Start_Row As Integer, _
    Start_Column As Integer
Dim Min_Fwd_CG As Double, _
    Max_Aft_CG As Double
Dim Cant_Punch_Tanks_Fwd As Boolean, _
    Cant_Punch_Tanks_Aft As Boolean, _
    Ignore_Case_Fwd As Boolean, _
    Ignore_Case_Aft As Boolean, _
    Punch_Tanks_on_4_or_6 As Boolean
Dim Current_Store_Name As String

'ID Mission Variables
Dim oWorkShtCst As Excel.Worksheet
Dim oRange As Excel.Range

Dim Expendables_FirstItem As String
Dim Expendables As Boolean
Dim Force_SA As Boolean

'Copy Worst Case Jett Variables

Dim oWorkShtIn As Excel.Worksheet
Dim oWorkShtOut As Excel.Worksheet

Dim oRangeIn As Excel.Range, _
    oRangeOut As Excel.Range, _
    oRangeRead As Excel.Range
Dim oCleanRange As Excel.Range, _
    oWorstAftWeight As Excel.Range, _
    oWorstFwdWeight As Excel.Range, _
    oAddGUMC As Excel.Range
Dim oMaxAftCGOut As Excel.Range, _
    oMaxFwdCGOut As Excel.Range
    
Dim Temp_Array2(6) As Variant, _
    Temp_Array1() As Variant
    
Sub Update_Form_F()

    'Purpose: Run subroutines when the Update Form F button is clicked
    Arrange_Stores
    Find_Worst_CG
    Copy_Worst_Case_Jett
    ID_Mission
    Fill_Remarks
    Switch_to_Form_F
    
End Sub

Function Percent_MAC(ByVal Lat_Mom As Double, Weight As Double, x As Double, y As Double)

    'Purpose: Compute the %MAC based on the total aircraft weight and longitudinal moment
    Internal_Var = (Lat_Mom * 100) / Weight
    Percent_MAC = (Internal_Var - x) / y

End Function

Function Chin_Station_Config(ByVal Chin_Stations As String) As String

    'Purpose: Convert a binary string to a text string representing the chin station config
    Select Case Chin_Stations
    
        Case Is = "00"
            Chin_Station_Config = "E"
        Case Is = "01"
            Chin_Station_Config = "R"
        Case Is = "10"
            Chin_Station_Config = "L"
        Case Is = "11"
            Chin_Station_Config = "LR"
    
    End Select
 
End Function

Function Applicable_Regulation(ByVal Model As String) As String

    Select Case Model
    
        Case Is = "F-16AM 10/15", Is = "F-16BM 10/15"
            Applicable_Regulation = "1F-16A-5-2"
        Case Is = "F-16C 25/30/32", Is = "F-16D 25/30/32"
            Applicable_Regulation = "1F-16C-5-2"
        Case Is = "F-16CM 40/42", Is = "F-16CM 50/52", Is = "F-16DM 40/42", Is = "F-16DM 50/52"
            Applicable_Regulation = "1F-16CM-5-2"
        
    End Select

End Function

Sub Fill_Remarks()

    Set oFormF = Worksheets("Form F")
    Set oCalcs = Worksheets("Calculations")
    Set oFormFRange = oFormF.Range("A45")
    
    oFormF.Range("A45:G59").Value = ""
    
    If oCalcs.Range("S3").Value = "00" Then
        Pod_Config = "NONE"
    Else
        Pod_Config = "ONE"
    End If
    
    If oCalcs.Range("Q3") = 3 Then
        Landing_Cat = "CATEGORY I"
    Else
        Landing_Cat = oCalcs.Range("R6")
    End If
    
    Counter = 0
    
    
    'Determine The Worst Case Nose Wheel Steering
    All_NWS = oCalcs.Range("CT3").Value & " " & oCalcs.Range("CT4").Value & " " & oCalcs.Range("CT5")
    
    If InStr(All_NWS, "Warning") <> 0 Then
        
        With oFormFRange
            .Offset(Counter, 0) = "NOSE WHEEL STEERING DISENGAGEMENT IS PROBABLE"
            .Offset(Counter + 1, 0) = "TAXI IN CONGESTED AREA IS NOT ADVISABLE"
            .Offset(Counter + 2, 0) = "TOWING MAY BE REQUIRED"
            Counter = Counter + 3
        End With
            
    ElseIf InStr(All_NWS, "Caution") <> 0 Then
    
        With oFormFRange
            .Offset(Counter, 0) = "NOSE WHEEL STEERING DISENGAGEMENT MAY OCCUR"
            .Offset(Counter + 1, 0) = "USE CAUTION WHEN TAXIING"
            Counter = Counter + 2
        End With
    
    End If
        
    
    With oFormFRange
        .Offset(Counter, 0).Value = "NOSE WHEEL STEERING " & oCalcs.Range("CT3").Value & " at TAKEOFF"
        .Offset(Counter + 1, 0).Value = "NOSE WHEEL STEERING " & oCalcs.Range("CT4").Value & " at LANDING"
        .Offset(Counter + 2, 0).Value = "NOSE WHEEL STEERING " & oCalcs.Range("CT5").Value & " at MOST AFT"
        .Offset(Counter + 3, 0).Value = "NOSE TIRE: 16VL027-1 18 PLY at 300-310 psi (" & oCalcs.Range("L3").Value & " lbs)"
        .Offset(Counter + 4, 0).Value = "GEAR RETRACTION: " & oCalcs.Range("K3")
        .Offset(Counter + 5, 0).Value = "INLET PODS: " & Pod_Config
        .Offset(Counter + 6, 0).Value = "FUEL LOADED at 6.8 LBS/GAL"
        .Offset(Counter + 7, 0).Value = "ESTIMATED LANDING FUEL: 1000 LBS (147 gallons)"
        .Offset(Counter + 8, 0).Value = oCalcs.Range("R6") & " LOADING at TAKEOFF"
        .Offset(Counter + 9, 0).Value = Landing_Cat & " LOADING at LANDING"
        .Offset(Counter + 10, 0).Value = oCalcs.Range("R6") & " LOADING at MOST AFT"
    End With
    
    If oCalcs.Range("H6").Value = "DUAL" And oCalcs.Range("AT11").Value = True Then
    
        With oFormFRange
            .Offset(Counter + 11, 0).Value = "*** FOR SOLO FLIGHT ONLY ***"
        End With
        
    End If
    
    
End Sub

Sub Arrange_Stores()

    Application.Calculation = xlCalculationManual

    'Purpose: Identify the loaded stores and sort them into jettisonable and non-jettisonable
    Set oWorkSht = Worksheets("Calculations")
    Set oWorkShtConfig = Worksheets("Configurator")
    Set oWorkShtStores = Worksheets("Stores")
    
    'CONSTANTS'
    Retained_Row_Count = 3
    Jett_Row_Count = 3
    Exp_Row_Count = 26
     
    Retained_Column = 35
    Jett_Column = 40
    Exp_Column = 40
    'END CONSTANTS
    
    oWorkSht.Range("AI3:AL22").Value = Null
    oWorkSht.Range("AN3:AQ22").Value = Null
    oWorkSht.Range("AN26:AQ45").Value = Null

    For Each C In oWorkSht.Range("AC3:AC28").Cells
        
        If C.Value <> 0 And C.Offset(0, 4) = 1 Then
            oWorkSht.Cells(Retained_Row_Count, Retained_Column) = C.Offset(29, 0).Value & " (STA " & C.Offset(0, -2) & ")"
            oWorkSht.Cells(Retained_Row_Count, Retained_Column + 1) = C.Offset(0, 1).Value
            oWorkSht.Cells(Retained_Row_Count, Retained_Column + 2) = C.Offset(0, 2).Value
            oWorkSht.Cells(Retained_Row_Count, Retained_Column + 3) = C.Offset(0, 3).Value
            Retained_Row_Count = Retained_Row_Count + 1
        End If
       
        If C.Value <> 0 And C.Offset(0, 4) = 3 Then
             oWorkSht.Cells(Exp_Row_Count, Exp_Column) = C.Offset(29, 0).Value & " (STA " & C.Offset(0, -2) & ")"
             oWorkSht.Cells(Exp_Row_Count, Exp_Column + 1) = C.Offset(0, 1).Value
             oWorkSht.Cells(Exp_Row_Count, Exp_Column + 2) = C.Offset(0, 2).Value
             oWorkSht.Cells(Exp_Row_Count, Exp_Column + 3) = C.Offset(0, 3).Value
             Exp_Row_Count = Exp_Row_Count + 1
             Exp_Lead_Text = ChrW(10003)
        End If
       
        If C.Value <> 0 And (C.Offset(0, 4) = 2 Or C.Offset(0, 4) = 3) Then
            oWorkSht.Cells(Jett_Row_Count, Jett_Column) = Exp_Lead_Text & C.Offset(29, 0).Value & " (STA " & C.Offset(0, -2) & ")"
            oWorkSht.Cells(Jett_Row_Count, Jett_Column + 1) = C.Offset(0, 1).Value
            oWorkSht.Cells(Jett_Row_Count, Jett_Column + 2) = C.Offset(0, 2).Value
            oWorkSht.Cells(Jett_Row_Count, Jett_Column + 3) = C.Offset(0, 3).Value
            Jett_Row_Count = Jett_Row_Count + 1
        End If
        
        Exp_Lead_Text = ""
       
    Next C

    'Chaff/Flare Copy
    If oWorkSht.Range("AA62").Value = True Then
    
        Exp_Lead_Text = ChrW(10003)
    
        oWorkSht.Cells(Retained_Row_Count, Retained_Column) = oWorkShtStores.Range("BD4")
        oWorkSht.Cells(Retained_Row_Count, Retained_Column + 1) = oWorkShtStores.Range("BE4")
        oWorkSht.Cells(Retained_Row_Count, Retained_Column + 2) = oWorkShtStores.Range("BF4")
        oWorkSht.Cells(Retained_Row_Count, Retained_Column + 3) = oWorkShtStores.Range("BG4")
        
        oWorkSht.Cells(Jett_Row_Count, Jett_Column) = Exp_Lead_Text & oWorkShtStores.Range("BD5")
        oWorkSht.Cells(Jett_Row_Count, Jett_Column + 1) = oWorkShtStores.Range("BE5")
        oWorkSht.Cells(Jett_Row_Count, Jett_Column + 2) = oWorkShtStores.Range("BF5")
        oWorkSht.Cells(Jett_Row_Count, Jett_Column + 3) = oWorkShtStores.Range("BG5")
        
        oWorkSht.Cells(Exp_Row_Count, Exp_Column) = oWorkShtStores.Range("BD5")
        oWorkSht.Cells(Exp_Row_Count, Exp_Column + 1) = oWorkShtStores.Range("BE5")
        oWorkSht.Cells(Exp_Row_Count, Exp_Column + 2) = oWorkShtStores.Range("BF5")
        oWorkSht.Cells(Exp_Row_Count, Exp_Column + 3) = oWorkShtStores.Range("BG5")
        
        Retained_Row_Count = Retained_Row_Count + 1
        Jett_Row_Count = Jett_Row_Count + 1
        Exp_Row_Count = Exp_Row_Count + 1
        
        Exp_Lead_Text = ""
        
    End If
    
    'Manual Stores Copy
    For Each C In oWorkShtConfig.Range("A52:A63").Cells
    
       If C.Text <> "" And C.Offset(0, 5) = 1 Then
           oWorkSht.Cells(Retained_Row_Count, Retained_Column) = C.Offset(0, 0).Value & " (STA " & C.Offset(0, 4) & ")"
           oWorkSht.Cells(Retained_Row_Count, Retained_Column + 1) = C.Offset(0, 1).Value
           oWorkSht.Cells(Retained_Row_Count, Retained_Column + 2) = C.Offset(0, 2).Value
           oWorkSht.Cells(Retained_Row_Count, Retained_Column + 3) = C.Offset(0, 3).Value
           Retained_Row_Count = Retained_Row_Count + 1
       End If
       
       If C.Text <> "" And C.Offset(0, 5) = 3 Then
           oWorkSht.Cells(Exp_Row_Count, Exp_Column) = C.Offset(0, 0).Value & " (STA " & C.Offset(0, 4) & ")"
           oWorkSht.Cells(Exp_Row_Count, Exp_Column + 1) = C.Offset(0, 1).Value
           oWorkSht.Cells(Exp_Row_Count, Exp_Column + 2) = C.Offset(0, 2).Value
           oWorkSht.Cells(Exp_Row_Count, Exp_Column + 3) = C.Offset(0, 3).Value
           Exp_Row_Count = Exp_Row_Count + 1
           Exp_Lead_Text = ChrW(10003)
       End If
       
       If C.Text <> "" And (C.Offset(0, 5) = 2 Or C.Offset(0, 5) = 3) Then
           oWorkSht.Cells(Jett_Row_Count, Jett_Column) = Exp_Lead_Text & C.Offset(0, 0).Value & " (STA " & C.Offset(0, 4) & ")"
           oWorkSht.Cells(Jett_Row_Count, Jett_Column + 1) = C.Offset(0, 1).Value
           oWorkSht.Cells(Jett_Row_Count, Jett_Column + 2) = C.Offset(0, 2).Value
           oWorkSht.Cells(Jett_Row_Count, Jett_Column + 3) = C.Offset(0, 3).Value
           Jett_Row_Count = Jett_Row_Count + 1
       End If
       
       Exp_Lead_Text = ""
       
    Next C

    Application.Calculation = xlCalculationAutomatic

End Sub

Sub Switch_to_Form_F()
    
    'Purpose: Change the user's focus to the Form F
    Worksheets("Form F").Activate

End Sub

Sub Switch_to_Configurator()
    
    'Purpose: Change the user's focus to the Form F
    Worksheets("Configurator").Activate

End Sub

Sub Find_Worst_CG()

    Application.Calculation = xlCalculationManual

    'Purpose: Evaluate all combinations of jettisonable stores to determine which moves the CG most forward and aft
    Set oWorkSht = Worksheets("Calculations")
    Set oWorkShtCst = Worksheets("Constants")
    
    oWorkSht.Range("BE3:BG4").Cells.Value = Null
    
    Total_Weight = 0
    Total_Lon_Moment = 0
    
    Counter = 3
    Column = 57
    Start_Row = 2
    Start_Column = 41
    
    Aft_WC_Weight = oWorkSht.Range("BJ3").Value
    Aft_WC_Mom = oWorkSht.Range("BK3").Value
    Fwd_WC_Weight = oWorkSht.Range("BL3").Value
    Fwd_WC_Mom = oWorkSht.Range("BM3").Value
    Datum_To_Leading_Edge = oWorkShtCst.Range("A3").Value
    Mac = oWorkShtCst.Range("B3").Value
    
    Cant_Punch_Tanks_Fwd = oWorkSht.Range("D6")
    Cant_Punch_Tanks_Aft = oWorkSht.Range("D8")
    
    'Initialize the output variables in case there are no Jettisionable stores
    Min_Fwd_CG = Percent_MAC(CDbl(Fwd_WC_Mom), CDbl(Fwd_WC_Weight), CDbl(Datum_To_Leading_Edge), CDbl(Mac))
    Max_Aft_CG = Percent_MAC(CDbl(Aft_WC_Mom), CDbl(Aft_WC_Weight), CDbl(Datum_To_Leading_Edge), CDbl(Mac))
    Max_Aft_Weight = 0
    Max_Aft_Mom = 0
    Max_Aft_Config = 0
    Min_Fwd_Weight = 0
    Min_Fwd_Mom = 0
    Min_Fwd_Config = 0
    Max_Aft_4_or_6_Tank_Punch = False
    Min_Fwd_4_or_6_Tank_Punch = False
    
    'Find out how many stores are Jettisonable
    
    Array_Length = WorksheetFunction.CountA(oWorkSht.Range("AN3:AN22").Cells)
                 
    'For each combination of stores, determine the total weight and total lon moment
    
    If Array_Length <> 0 Then
    
        For Each var In Binary_Strings(Array_Length)
            Ignore_Case_Fwd = False
            Ignore_Case_Aft = False
            Punch_Tanks_on_4_or_6 = False
            'Sum the weights and moments from the combination of jettisoned stores
            For var_char = 1 To Len(var)
                If Mid(var, var_char, 1) = 1 Then
                    Current_Store_Name = oWorkSht.Cells(Start_Row + var_char, Start_Column - 1)
                    'Abandon assessment of this jettision possibility if it violates the fuel assumption
                    If InStr(Current_Store_Name, "370 TANK") <> 0 _
                        And Cant_Punch_Tanks_Fwd = True Then
                        Ignore_Case_Fwd = True
                    End If
                    If InStr(Current_Store_Name, "370 TANK") <> 0 _
                        And Cant_Punch_Tanks_Aft = True Then
                        Ignore_Case_Aft = True
                    End If
                    If InStr(Current_Store_Name, "370 TANK") <> 0 And InStr(Current_Store_Name, "STA 4") <> 0 Then
                        Punch_Tanks_on_4_or_6 = True
                    End If
                    Total_Weight = Total_Weight + oWorkSht.Cells(Start_Row + var_char, Start_Column).Value
                    Total_Lon_Moment = Total_Lon_Moment + oWorkSht.Cells(Start_Row + var_char, Start_Column + 1).Value
                End If
            Next var_char
        
            'Calculate the %MACs for both Aft and Forward Cases
            Temp_Aft_CG = Percent_MAC(Aft_WC_Mom - Total_Lon_Moment, Aft_WC_Weight - Total_Weight, CDbl(Datum_To_Leading_Edge), CDbl(Mac))
            Temp_Fwd_CG = Percent_MAC(Fwd_WC_Mom - Total_Lon_Moment, Fwd_WC_Weight - Total_Weight, CDbl(Datum_To_Leading_Edge), CDbl(Mac))
        
            'If the %MAC calculated is the worst case so far, replace the previous worst case
            If Temp_Aft_CG > Max_Aft_CG And Ignore_Case_Aft <> True Then
                Max_Aft_CG = Temp_Aft_CG
                Max_Aft_Weight = -Total_Weight
                Max_Aft_Mom = -Total_Lon_Moment
                Max_Aft_Config = var
                Max_Aft_4_or_6_Tank_Punch = Punch_Tanks_on_4_or_6
            End If
                
            If Temp_Fwd_CG < Min_Fwd_CG And Ignore_Case_Fwd <> True Then
                Min_Fwd_CG = Temp_Fwd_CG
                Min_Fwd_Weight = -Total_Weight
                Min_Fwd_Mom = -Total_Lon_Moment
                Min_Fwd_Config = var
                Min_Fwd_4_or_6_Tank_Punch = Punch_Tanks_on_4_or_6
            End If
                
            Total_Weight = 0
            Total_Lon_Moment = 0
            
            Counter = Counter + 1
        
        Next var
    
    End If
        
    'Update the Calculations worksheet with the new values
    oWorkSht.Range("BE3").Cells.Value = Min_Fwd_Config
    oWorkSht.Range("BF3").Cells.Value = Min_Fwd_Weight
    oWorkSht.Range("BG3").Cells.Value = Min_Fwd_Mom
    oWorkSht.Range("BH3").Cells.Value = Min_Fwd_CG
    oWorkSht.Range("BI3").Cells.Value = Min_Fwd_4_or_6_Tank_Punch
    oWorkSht.Range("BE4").Cells.Value = Max_Aft_Config
    oWorkSht.Range("BF4").Cells.Value = Max_Aft_Weight
    oWorkSht.Range("BG4").Cells.Value = Max_Aft_Mom
    oWorkSht.Range("BH4").Cells.Value = Max_Aft_CG
    oWorkSht.Range("BI4").Cells.Value = Max_Aft_4_or_6_Tank_Punch
    Application.Calculation = xlCalculationAutomatic
    
End Sub

    
Function Binary_Strings(ByVal input_value As Integer) As Collection

    'Purpose: Convert a number into all combinations of that number of bits in string format
    Dim var As New Collection

    Collection_Length = 2 ^ input_value - 1
    Counter = Collection_Length
    
    If input_value <> 0 Then
        For i = 1 To Collection_Length + 1
            current_string = WorksheetFunction.Dec2Bin(Counter, input_value)
            var.Add (current_string)
            Counter = Counter - 1
        Next i
        
    Set Binary_Strings = var
    
    End If
    
End Function
    
Sub ID_Mission()

    'Determine the mission of the aircraft based on loading for use with the taxi/takeoff fuel lookup table
    Set oWorkSht = Worksheets("Calculations")
    Set oRange = oWorkSht.Range("BZ3")
    
    Station_3a = oWorkSht.Range("AC8").Value
    Station_3b = oWorkSht.Range("AC10").Value
    Station_4a = oWorkSht.Range("AC11").Value
    Station_4b = oWorkSht.Range("AC13").Value
    Station_5 = oWorkSht.Range("AC15").Value
    Station_6a = oWorkSht.Range("AC18").Value
    Station_6b = oWorkSht.Range("AC20").Value
    Station_7a = oWorkSht.Range("AC21").Value
    Station_7b = oWorkSht.Range("AC23").Value
    
    Force_SA = oWorkSht.Range("BY5").Value
            
    Expendables_FirstItem = oWorkSht.Range("AN26").Value
    
    Expendables = False
    
    If Expendables_FirstItem <> "" Or Force_SA Then
        
        If Expendables_FirstItem <> "Chaff/Flare (Expendable)" Then
        
            Expendables = True
            
        End If
        
    End If
        
    If (InStr(1, Station_4a, "TANK", vbTextCompare) <> 0 Or _
        InStr(1, Station_4b, "TANK", vbTextCompare) <> 0) And _
        InStr(1, Station_5, "TANK", vbTextCompare) <> 0 And _
        (InStr(1, Station_6a, "TANK", vbTextCompare) <> 0 Or _
        InStr(1, Station_6b, "TANK", vbTextCompare) <> 0) Then
        
        If Expendables Then
        
            oRange.Value = "Air-to-Surface"
            Exit Sub
        
        End If
        
        oRange.Value = "3 Tanks"
        Exit Sub
        
    End If
    
    If (InStr(1, Station_4a, "TANK", vbTextCompare) <> 0 Or _
       InStr(1, Station_4b, "TANK", vbTextCompare) <> 0) And _
        (InStr(1, Station_6a, "TANK", vbTextCompare) <> 0 Or _
        InStr(1, Station_6b, "TANK", vbTextCompare) <> 0) Then
             
        If Expendables Then
        
            oRange.Value = "Air-to-Surface"
            Exit Sub
        
        End If
        
        oRange.Value = "Wing Tanks"
        Exit Sub
        
    End If
    
    If InStr(1, Station_5, "TANK", vbTextCompare) <> 0 Then
            
        If Expendables Then
        
            oRange.Value = "Air-to-Surface"
            Exit Sub
        
        End If
        
        oRange.Value = "Centerline"
        Exit Sub
    
    End If
    
    If Expendables Then
    
        oRange.Value = "Air-to-Surface"
        Exit Sub
    End If
    
    oRange.Value = "Clean"
    
End Sub

Sub Print_Form_F()

    On_Stores_Dropdown_Click
    Worksheets("Form F").PrintOut

End Sub

Sub Copy_Worst_Case_Jett()

    Application.Calculation = xlCalculationManual

    'Using the binary codes produced by Find_Worst_CG, copy the stores with weight and lon. moment data to the Form F
    Set oWorkShtIn = Worksheets("Calculations")
    Set oWorkShtOut = Worksheets("Form F")
    Set oRangeIn = oWorkShtIn.Range("BE3:BE4")
    Set oRangeRead = oWorkShtIn.Range("AN3:AQ3")
    Set oRangeOut = oWorkShtOut.Range("A12:G12")
    Set oWorstFwdWeight = oWorkShtIn.Range("BO5")
    Set oWorstAftWeight = oWorkShtIn.Range("BO7")
    Set oAddGUMC = oWorkShtIn.Range("Y6")
    Set oMaxFwdCGOut = oWorkShtOut.Range("L70")
    Set oMaxAftCGOut = oWorkShtOut.Range("L71")
    
    Set oCleanRange = oWorkShtOut.Range("A12:G43")
        
    With oCleanRange
        .Value = Null
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        .Font.Bold = False
        .UnMerge
        oWorkShtOut.Range(.Cells(1, 1), "C43").HorizontalAlignment = xlLeft
        oWorkShtOut.Range(.Cells(1, 4), "G43").HorizontalAlignment = xlRight
        oWorkShtOut.Range(.Cells(1, 7), "G43").NumberFormat = "0.0%"
    End With

    Worst_Case_Fwd = oRangeIn.Value2(1, 1)
    Worst_Case_Aft = oRangeIn.Value2(2, 1)
        
    Outrow = 0
    
    'Add the 95th % Pilots Line if Dual with No Backseater
    If oWorkShtIn.Range("H6").Value = "DUAL" And oWorkShtIn.Range("AT11").Value = True Then
    
        Temp_Array2(0) = "95th % Pilots"
        Temp_Array2(3) = "=Calculations!AT14"
        Temp_Array2(4) = "=Calculations!AU14"
        Temp_Array2(6) = "=Percent_MAC(SUM(E8:E" & Outrow + 12 & "),SUM(D8:D" & Outrow + 12 & "),Constants!$A$3,Constants!$B$3)"
    
        With oRangeOut.Offset(Outrow, 0)
            .Value2 = Temp_Array2
            oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
            oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
    
        Outrow = Outrow + 1
    
    End If
    
    'Enter all the worst case forward jettisoned stores
    For var_char = 1 To Len(Worst_Case_Fwd)
        
        If Mid(Worst_Case_Fwd, var_char, 1) = "1" Then
        
            Temp_Array1() = oRangeRead.Offset(var_char - 1, 0).Value2
            Temp_Array2(0) = Temp_Array1(1, 1)
            Temp_Array2(1) = ""
            Temp_Array2(2) = ""
            Temp_Array2(3) = -Temp_Array1(1, 2)
            Temp_Array2(4) = -Temp_Array1(1, 3)
            Temp_Array2(5) = ""
            Temp_Array2(6) = "=Percent_MAC(SUM(E8:E" & Outrow + 12 & "),SUM(D8:D" & Outrow + 12 & "),Constants!$A$3,Constants!$B$3)"
            
            With oRangeOut.Offset(Outrow, 0)
                oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
                oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
                .Value2 = Temp_Array2
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
            
            Outrow = Outrow + 1
            
        End If
        
    Next var_char
    
    'Add most forward fuel condition line 1
    Temp_Array2(0) = "MOST FORWARD CONDITION"
    Temp_Array2(3) = "=SUM(D8:D" & (Outrow + 11) & ")"
    Temp_Array2(4) = "=SUM(E8:E" & (Outrow + 11) & ")"
    Temp_Array2(6) = ""
        
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oWorstFwdWeight.Value = .Cells(1, 4).Value
        .Cells(1, 1).Font.Bold = True
        Range(.Cells(1, 1), .Cells(1, 3)).Merge
        Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
        
    Outrow = Outrow + 1
    
    'Add most forward fuel condition line 2
    Temp_Array2(0) = "MOST FWD CG LIMIT vs ACTUAL"
    Temp_Array2(3) = ""
    Temp_Array2(4) = "=CONCATENATE(TEXT(Configurator!L46," & Chr$(34) & _
        "0.0%" & Chr$(34) & ")," & Chr$(34) & "   < " & Chr$(34) & ")"
    Temp_Array2(6) = "=Percent_MAC(E" & (Outrow + 11) & ",D" & (Outrow + 11) & ",Constants!$A$3,Constants!$B$3)"
    
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oMaxFwdCGOut.Value = .Cells(1, 7).Value
        Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
        
    Outrow = Outrow + 1
   
    With oRangeOut.Offset(Outrow, 0)
        .Cells(1, 1).Value = "M O S T  A F T"
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    Outrow = Outrow + 1
    
    'Add Ramp Weight Line
    Temp_Array2(0) = "RAMP WEIGHT"
    Temp_Array2(3) = "=K54"
    Temp_Array2(4) = "=L54"
    Temp_Array2(6) = "=Percent_MAC(E" & (Outrow + 12) & ",D" & (Outrow + 12) & ",Constants!$A$3,Constants!$B$3)"
    
    New_Sum_Start = Outrow + 12
    
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        .Cells(1, 1).Font.Bold = True
        oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    Outrow = Outrow + 1
    
    'Add Remove All Fuel
    Temp_Array2(0) = "REMOVE ALL FUEL"
    Temp_Array2(3) = "=-SUM(K$52:K$53)"
    Temp_Array2(4) = "=-SUM(L$52:L$53)"
    Temp_Array2(6) = "=Percent_MAC(SUM(E" & New_Sum_Start & ":E" & (Outrow + 12) & "),SUM(D" & _
        New_Sum_Start & ":D" & (Outrow + 12) & "),Constants!$A$3,Constants!$B$3)"
    
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        .Cells(1, 1).Font.Bold = True
        oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        
    End With
    
    Outrow = Outrow + 1
    
    'Add Most AFT Fuel Condition
    Temp_Array2(0) = "Most AFT Fuel Condition"
    Temp_Array2(3) = "=Calculations!O3"
    Temp_Array2(4) = "=Calculations!P3"
    Temp_Array2(6) = "=Percent_MAC(SUM(E" & New_Sum_Start & ":E" & (Outrow + 12) & "),SUM(D" & _
        New_Sum_Start & ":D" & (Outrow + 12) & "),Constants!$A$3,Constants!$B$3)"
        
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    Outrow = Outrow + 1
    
    'Add GUMC
    If oAddGUMC.Value = "True" Then
        Temp_Array2(0) = "Gear-Up Moment (GUMC)"
        Temp_Array2(3) = ""
        Temp_Array2(4) = "=Calculations!K3"
        Temp_Array2(6) = "=Percent_MAC(SUM(E" & New_Sum_Start & ":E" & (Outrow + 12) & "),SUM(D" & _
            New_Sum_Start & ":D" & (Outrow + 12) & "),Constants!$A$3,Constants!$B$3)"
        
        With oRangeOut.Offset(Outrow, 0)
            .Value2 = Temp_Array2
            oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
            oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
        Outrow = Outrow + 1
        
    End If
    
    'Add 5th % Pilots
    Temp_Array2(0) = "5th % Pilots"
    Temp_Array2(3) = "=Calculations!AU9"
    Temp_Array2(4) = "=Calculations!AV9"
    Temp_Array2(6) = "=Percent_MAC(SUM(E" & New_Sum_Start & ":E" & (Outrow + 12) & "),SUM(D" & _
        New_Sum_Start & ":D" & (Outrow + 12) & "),Constants!$A$3,Constants!$B$3)"
    
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    Outrow = Outrow + 1
    
    'Add worst case aft jettisoned stores
    For var_char = 1 To Len(Worst_Case_Aft)
 
        If Mid(Worst_Case_Aft, var_char, 1) = "1" Then
        
            Temp_Array1() = oRangeRead.Offset(var_char - 1, 0).Value2
            Temp_Array2(0) = Temp_Array1(1, 1)
            Temp_Array2(1) = ""
            Temp_Array2(2) = ""
            Temp_Array2(3) = -Temp_Array1(1, 2)
            Temp_Array2(4) = -Temp_Array1(1, 3)
            Temp_Array2(5) = ""
            Temp_Array2(6) = "=Percent_MAC(SUM(E" & New_Sum_Start & " :E" & Outrow + 12 & "),SUM(D" & New_Sum_Start & ":D" & _
                Outrow + 12 & "),Constants!$A$3,Constants!$B$3)"
            
            With oRangeOut.Offset(Outrow, 0)
                oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
                oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
                .Value2 = Temp_Array2
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
            
            Outrow = Outrow + 1
            
        End If
        
    Next var_char
    
    'Add most forward fuel condition line 1
    Temp_Array2(0) = "MOST AFT CONDITION"
    Temp_Array2(3) = "=SUM(D" & New_Sum_Start & ":D" & (Outrow + 11) & ")"
    Temp_Array2(4) = "=SUM(E" & New_Sum_Start & ":E" & (Outrow + 11) & ")"
    Temp_Array2(6) = ""
        
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oWorstAftWeight.Value = .Cells(1, 4).Value
        .Cells(1, 1).Font.Bold = True
        oWorkShtOut.Range(.Cells(1, 1), .Cells(1, 3)).Merge
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
        
    Outrow = Outrow + 1
    
    'Add most forward fuel condition line 2
    Temp_Array2(0) = "MOST AFT CG LIMIT vs ACTUAL"
    Temp_Array2(3) = ""
    Temp_Array2(4) = "=CONCATENATE(TEXT(Configurator!O46," & Chr$(34) & _
        "0.0%" & Chr$(34) & ")," & Chr$(34) & "   > " & Chr$(34) & ")"
    Temp_Array2(6) = "=Percent_MAC(E" & (Outrow + 11) & ",D" & (Outrow + 11) & ",Constants!$A$3,Constants!$B$3)"
    
    With oRangeOut.Offset(Outrow, 0)
        .Value2 = Temp_Array2
        oMaxAftCGOut.Value = .Cells(1, 7).Value
        oWorkShtOut.Range(.Cells(1, 5), .Cells(1, 6)).Merge
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    
    Outrow = Outrow + 1
    
    With oRangeOut.Offset(Outrow, 0)
        oWorkShtOut.Range(.Cells(1, 1), "C43").Borders(xlEdgeRight).LineStyle = xlContinuous
        oWorkShtOut.Range(.Cells(1, 4), "D43").Borders(xlEdgeRight).LineStyle = xlContinuous
        oWorkShtOut.Range(.Cells(1, 5), "F43").Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    
    Application.Calculation = xlCalculationAutomatic
        
End Sub
