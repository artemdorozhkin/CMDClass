Attribute VB_Name = "indexCmdIwi"
'@Folder "CmdIwiProject.src"
Option Explicit

Public Static Function cmd(Optional Byval CreateNew As Boolean = False) As CmdIwi
    Dim StaticCMD As CmdIwi

    If StaticCMD Is Nothing Or CreateNew Then
        Set StaticCMD = New CmdIwi
    End If

    Set cmd = StaticCMD
End Function