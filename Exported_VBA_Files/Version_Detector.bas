Attribute VB_Name = "Version_Detector"
Function Sheet_Version() As String
    
    var = Split(ActiveWorkbook.Name, "v")
    Sheet_Version = Left(var(1), 3)

End Function
