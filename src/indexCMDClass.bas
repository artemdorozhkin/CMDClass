'@Folder "CMDClassProject.src"
Option Explicit
 
Public Static Function cmd(Optional ByVal CreateNew As Boolean = False) As CMDClass
    Dim StaticCMD As CMDClass

    If StaticCMD Is Nothing Or CreateNew Then
        Set StaticCMD = New CMDClass
    End If

    Set cmd = StaticCMD
End Function