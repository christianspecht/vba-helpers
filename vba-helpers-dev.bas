Option Compare Database
Option Explicit

Const vbahelpersdevfilename As String = "vba-helpers-dev.bas"
Const vbahelpersdevmodulename As String = "VBAHelpersDev"
Const vbahelperstestfilename As String = "vba-helpers-tests.bas"
Const vbahelperstestmodulename As String = "VBAHelpersTests"

Public Sub VBAHelpers_Export()
    'Exports all modules to the current directory (for source control) and increases the version number in the VBA Helpers module.

    Const versionstring As String = "'# Version "
    Dim exportfile As String
    
    'export VBA Helpers
    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelpersfilename)
    Application.SaveAsText acModule, vbahelpersmodulename, exportfile

    'set version number
    Dim lines1() As String
    Dim lines2() As String
    Dim i As Long

    lines1 = File_ReadAllLines(exportfile)
    ReDim lines2(UBound(lines1))
    For i = 0 To UBound(lines1)
        If String_StartsWith(lines1(i), versionstring) Then
            lines2(i) = versionstring & format(Now, "yyyymmdd.hhmmss")
        Else
            lines2(i) = lines1(i)
        End If
    Next

    File_WriteAllLines exportfile, lines2

    'export tests
    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelperstestfilename)
    Application.SaveAsText acModule, vbahelperstestmodulename, exportfile

    'export dev functions
    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelpersdevfilename)
    Application.SaveAsText acModule, vbahelpersdevmodulename, exportfile

End Sub

Public Sub VBAHelpers_Import()
    'Imports all VBA Helpers modules from the current directory.

    Dim exportfile As String
    Dim message As String

    'import tests
    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelperstestfilename)
    If Dir(exportfile) = "" Then
        message = String_Format("Couldn't find test class:{0}{1}", vbCrLf, exportfile)
        MsgBox message, vbCritical
        Exit Sub
    End If
    
    Application.LoadFromText acModule, vbahelperstestmodulename, exportfile


    'import dev functions
    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelpersdevfilename)
    If Dir(exportfile) = "" Then
        message = String_Format("Couldn't find dev functions:{0}{1}", vbCrLf, exportfile)
        MsgBox message, vbCritical
        Exit Sub
    End If
    
    Application.LoadFromText acModule, vbahelpersdevmodulename, exportfile
    

    'import VBA Helpers (use update function from actual VBA Helpers module)
    VBAHelpers_Update

End Sub