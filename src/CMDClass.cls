'@Folder "CMDClassProject.src"
Option Explicit

Private Type TCMDClass
    Shell As Object
    CurrentDir As String
    Value As String
End Type

Private this As TCMDClass

Private Sub Class_Initialize()
    Set this.Shell = CreateObject("WScript.Shell")
End Sub

'@DefaultMember
Public Property Get Value() As String
    Value = this.Value
End Property

Public Function Exec(ByVal Command As String) As CMDClass
    Const CMDCommand As String = "cmd /c "
    SetValue this.Shell.Exec(CMDCommand & Command).StdOut.ReadAll()
    Debug.Print this.Value

    Set Exec = Me
End Function

Public Function CD(Optional ByVal Path As String = ".") As CMDClass
    If Left(Path, 1) = "%" And Right(Path, 1) = "%" Then
        Path = this.Shell.ExpandEnvironmentStrings(Path)
    End If
    this.Shell.CurrentDirectory = Path
    this.Value = this.Shell.CurrentDirectory
    Set CD = Me
End Function

Public Function MkDir(ByVal FolderName As String) As CMDClass
    Exec "mkdir """ & FolderName & """"
    Set MkDir = Me
End Function

Public Function RmDir(ByVal FolderName As String) As CMDClass
    Exec "rmdir """ & FolderName & """"
    Set RmDir = Me
End Function

Public Function Dir() As CMDClass
    Exec "dir"
    Set Dir = Me
End Function

Public Function LS() As CMDClass
    Set LS = Me.Dir()
End Function

Public Function Copy(ByVal SourceFile As String, ByVal DestinationFolder As String) As CMDClass
    Exec "copy """ & SourceFile & """ """ & DestinationFolder & """"
    Set Copy = Me
End Function

Public Function Del(ByVal FileName As String) As CMDClass
    Exec "del """ & FileName & """"
    Set Del = Me
End Function

Public Function Move(ByVal SourceFile As String, ByVal DestinationFolder As String) As CMDClass
    Exec "move """ & SourceFile & """ """ & DestinationFolder & """"
    Set Move = Me
End Function

Public Function Rename(ByVal OldName As String, ByVal NewName As String) As CMDClass
    Exec "ren """ & OldName & """ """ & NewName & """"
    Set Rename = Me
End Function

Public Function TypeFile(ByVal FileName As String) As CMDClass
    Exec "type """ & FileName & """"
    Set TypeFile = Me
End Function

Public Function Touch(ByVal FileName As String, Optional ByVal Text As String = "NULL") As CMDClass
    If Text <> "NULL" Then
        Exec "echo " & Text & ">" & FileName
    Else
        Exec Text & ">" & FileName
    End If
    Set Touch = Me
End Function

Public Function Echo(Optional ByVal Text As String) As CMDClass
    Exec "echo " & Text
    Set Echo = Me
End Function

Public Function Start(ByVal Source As String) As CMDClass
    Exec "start " & Source
    Set Start = Me
End Function

Public Function TaskList() As CMDClass
    Exec "tasklist"
    Set TaskList = Me
End Function

Public Function SystemInfo() As CMDClass
    Exec "systeminfo "
    Set SystemInfo = Me
End Function

Public Function Regedit() As CMDClass
    Exec "regedit "
    Set Regedit = Me
End Function

Public Function Ping(ByVal Target As String) As CMDClass
    Exec "ping " & Target
    Set Ping = Me
End Function

Public Function IpConfig() As CMDClass
    Exec "ipconfig"
    Set IpConfig = Me
End Function

Public Function Netstat() As CMDClass
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