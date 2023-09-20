' Module QuickConstructors
    Function NewFSO()
        Set NewFSO = CreateObject("Scripting.FileSystemObject")
    End Function

    Function NewFolder(Path)
        Set NewFolder = NewFSO().GetFolder(Path)
    End Function

    Function NewFile(Path)
        Set NewFile = NewFSO().GetFile(Path)
    End Function

    Function NewExcel()
        Set NewExcel = CreateObject("Excel.Application")
    End Function
' End Module


' Module PathService
    Function GetRootPath()
        Dim FSO: Set FSO = NewFSO()
        GetRootPath = FSO.GetParentFolderName(WScript.ScriptFullName)

        Set FSO = Nothing
    End Function

    Function GetAbsolutePath(relativePath)
        Dim FSP: Set FSO = NewFSO()
        Dim rootPath: rootPath = GetRootPath()
        GetAbsolutePath = FSO.BuildPath(rootPath, relativePath)

        Set FSO = Nothing
    End Function
' End Module


' Module ModuleManager
    Sub ImportModulesFromFolder(Book, Folder)
        If Folder.SubFolders.Count > 0 Then 
            Dim SubFolder
            For Each SubFolder In Folder.SubFolders
                ImportModulesFromFolder Book, SubFolder
            Next
        End If

        Dim File
        For Each File in Folder.Files
            If IsVBAModule(File) Then
                Book.VBProject.VBComponents.Import File.Path 
            End If
        Next
    End Sub

    Function IsVBAModule(File)
        Dim FSO: Set FSO = NewFSO()
        Dim FileExtension: FileExtension = FSO.GetExtensionName(File.Path)

        Dim Extensions: Extensions = Array("bas", "cls", "frm", "doccls")
        Dim Extension
        For Each Extension In Extensions
            If StrComp(FileExtension, Extension, vbTextCompare) = 0 Then
                IsVBAModule = True
                Exit Function
            End If
        Next

        Set FSO = Nothing
    End Function

    Sub ImportPackageModule(Book)
        Dim FSO: Set FSO = NewFSO()
        Dim rootPath: rootPath = GetRootPath()
        Dim packagePath: packagePath = FSO.BuildPath(rootPath, "ThisProjectPackage.bas")

        If FSO.FileExists(packagePath) Then
            Book.VBProject.VBComponents.Import packagePath 
        End If

        Set FSO = Nothing
    End Sub
' End Module


' Module Main
    Sub Main()
        Dim Excel: Set Excel = NewExcel()
        Dim Book: Set Book = Excel.Workbooks.Add

        Dim srcPath: srcPath  = GetAbsolutePath("src")
        Dim FSO: Set FSO = NewFSO()
        Dim SrcFolder: Set SrcFolder = FSO.GetFolder(srcPath)

        ImportModulesFromFolder Book, SrcFolder
        ImportPackageModule Book
        Dim rootFolderName: rootFolderName = NewFolder(GetRootPath()).Name

        Const XLSBFormat = 50
        Dim savePath: savePath = FSO.BuildPath(GetRootPath(), rootFolderName & ".xlsb")
        Book.SaveAs savePath, XLSBFormat

        Book.Close False
        Excel.Quit

        Set Book = Nothing
        Set Excel = Nothing
    End Sub

    Main()
' End Module
