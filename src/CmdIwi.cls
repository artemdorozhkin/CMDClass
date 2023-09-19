VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdIwi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "CmdIwiProject.src"
Option Explicit

Private Type TCmdIwi
    Shell As Object
    CurrentDir As String
    Value As String
End Type

Private this As TCmdIwi

Private Sub Class_Initialize()
    Set this.Shell = CreateObject("WScript.Shell")
End Sub

'@DefaultMember
Public Property Get Value() As String
    Attribute Value.VB_UserMemId = 0
    Value = this.Value
End Property

Public Function Exec(ByVal Command As String) As CmdIwi
    Const CMDCommand As String = "cmd /c "
    SetValue this.Shell.Exec(CMDCommand & Command).StdOut.ReadAll()
    Debug.Print this.Value

    Set Exec = Me
End Function

Public Function CD(Optional ByVal Path As String = ".") As CmdIwi
    If Left(Path, 1) = "%" And Right(Path, 1) = "%" Then
        Path = this.Shell.ExpandEnvironmentStrings(Path)
    End If
    this.Shell.CurrentDirectory = Path
    this.Value = this.Shell.CurrentDirectory
    Set CD = Me
End Function

Public Function MkDir(ByVal FolderName As String) As CmdIwi
    Exec "mkdir """ & FolderName & """"
    Set MkDir = Me
End Function

Public Function RmDir(ByVal FolderName As String) As CmdIwi
    Exec "rmdir """ & FolderName & """"
    Set RmDir = Me
End Function

Public Function Dir() As CmdIwi
    Exec "dir"
    Set Dir = Me
End Function

Public Function LS() As CmdIwi
    Set LS = Me.Dir()
End Function

Public Function Copy(ByVal SourceFile As String, ByVal DestinationFolder As String) As CmdIwi
    Exec "copy """ & SourceFile & """ """ & DestinationFolder & """"
    Set Copy = Me
End Function

Public Function Del(ByVal FileName As String) As CmdIwi
    Exec "del """ & FileName & """"
    Set Del = Me
End Function

Public Function Move(ByVal SourceFile As String, ByVal DestinationFolder As String) As CmdIwi
    Exec "move """ & SourceFile & """ """ & DestinationFolder & """"
    Set Move = Me
End Function

Public Function Rename(ByVal OldName As String, ByVal NewName As String) As CmdIwi
    Exec "ren """ & OldName & """ """ & NewName & """"
    Set Rename = Me
End Function

Public Function TypeFile(ByVal FileName As String) As CmdIwi
    Exec "type """ & FileName & """"
    Set TypeFile = Me
End Function

Public Function Touch(ByVal FileName As String, Optional ByVal Text As String = "NULL") As CmdIwi
    If Text <> "NULL" Then
        Exec "echo " & Text & ">" & FileName
    Else
        Exec Text & ">" & FileName
    End If
    Set Touch = Me
End Function

Public Function Echo(Optional ByVal Text As String) As CmdIwi
    Exec "echo " & Text
    Set Echo = Me
End Function

Public Function Start(ByVal Source As String) As CmdIwi
    Exec "start " & Source
    Set Start = Me
End Function

Public Function TaskList() As CmdIwi
    Exec "tasklist"
    Set TaskList = Me
End Function

Public Function SystemInfo() As CmdIwi
    Exec "systeminfo "
    Set SystemInfo = Me
End Function

Public Function Regedit() As CmdIwi
    Exec "regedit "
    Set Regedit = Me
End Function

Public Function Ping(ByVal Target As String) As CmdIwi
    Exec "ping " & Target
    Set Ping = Me
End Function

Public Function IpConfig() As CmdIwi
    Exec "ipconfig"
    Set IpConfig = Me
End Function

Public Function Netstat() As CmdIwi
    Exec "netstat"
    Set Netstat = Me
End Function

Public Sub Clear()
    Application.SendKeys "^a^g{DEL}"
End Sub

Public Sub Cls()
    Me.Clear
End Sub

Private Sub SetValue(ByVal Value As String)
    Dim Stream As Object: Set Stream = CreateObject("ADODB.Stream")

    With Stream
        .Charset = "Windows-1251"
        .Open
        .WriteText Value
        .Position = 0
        .Charset = "cp866"
        this.Value = .ReadText
    End With
End Sub