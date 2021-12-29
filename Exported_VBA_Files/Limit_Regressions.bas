Attribute VB_Name = "Limit_Regressions"
Function Asymmetric_Limit(CG As Double, Weight As Double, Block As Boolean) As Double

    Select Case Block
    
        Case Is = False
    
            Asymmetric_Limit = (-13173.5 * CG) - (0.418875 * Weight) + 20425.7
        
        Case Is = True

            Asymmetric_Limit = (-14990.7 * CG) - (0.416894 * Weight) + 24497.7
        
        End Select
 
End Function


Function Nose_Wheel_Zone(CG As Double, Weight As Double, Block As String) As String
  
    Select Case Block
        
        Case Is = "F-16AM 10/15", Is = "F-16BM 10/15", Is = "F-16C 25/30/32", Is = "F-16D 25/30/32"
    
            Out_Of_Limits = 28.0197 * Exp(15.5307 * CG)
            Warning = 323.774 * Exp(9.5894 * CG)
            Caution = 353.884 * Exp(10.5888 * CG)
                    
        Case Is = "F-16CM 40/42", Is = "F-16CM 50/52", Is = "F-16DM 40/42", Is = "F-16DM 50/52"

            Out_Of_Limits = 1.95247 * Exp(21.5901 * CG)
            Warning = 289.398 * Exp(10.7338 * CG)
            Caution = 627.374 * Exp(9.2721 * CG)
            
        End Select

        If Weight < Out_Of_Limits Then
            Nose_Wheel_Zone = "OUT OF LIMITS"
        ElseIf Weight < Warning Then
            Nose_Wheel_Zone = "WARNING"
        ElseIf Weight < Caution Then
            Nose_Wheel_Zone = "CAUTION"
        Else
            Nose_Wheel_Zone = "NORMAL"
        End If
                        
End Function
