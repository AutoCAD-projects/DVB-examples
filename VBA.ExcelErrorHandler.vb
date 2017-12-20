Public Sub DisplayMessage( _
ByVal procedure As String, _
ByVal module As String, _
errNbr As Double, _
errDes As String, _
Optional ByVal errLine As Variant = 0, _
Optional ByVal title As String = "Unexpected Error")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:  Global error message for all procedures
' Example:  Call DisplayMessage("Module", "Procedure", 101, "descr", 1, "Error Description")
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim msg As String

    msg = "Contact your system administrator."
    msg = msg & vbCrLf & "Module: " & module
    msg = msg & vbCrLf & "Procedure: " & procedure
    msg = msg & IIf(errLine = 0, "", vbCrLf & "Error Line: " & errLine)
    msg = msg & vbCrLf & "Error #: " & errNbr
    msg = msg & vbCrLf & "Error Description: " & errDes
    MsgBox msg, vbCritical, title

End Sub
