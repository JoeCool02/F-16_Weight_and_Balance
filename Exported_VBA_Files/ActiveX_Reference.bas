Attribute VB_Name = "ActiveX_Reference"
Sub Add_ActiveX_Reference()
     'Macro purpose:  To add a reference to the project using the GUID for the
     'reference library
     
    Dim strGUID As String, theRef As Variant, i As Long
     
     'Update the GUID you need below.
    strGUID = "{2A75196C-D9EB-4129-B803-931327F72D5C}"
     
     'Set to continue in case of error
    On Error GoTo Reference_Error
     
    ThisWorkbook.VBProject.References.AddFromGuid _
    GUID:=strGUID, Major:=2, Minor:=5
     
Reference_Error:
    
End Sub


