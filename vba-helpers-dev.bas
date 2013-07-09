Option Compare Database
Option Explicit

Const vbahelpersdevfilename_vbah As String = "vba-helpers-dev.bas"
Const vbahelpersdevmodulename_vbah As String = "VBAHelpersDev"
Const vbahelperstestfilename_vbah As String = "vba-helpers-tests.bas"
Const vbahelperstestmodulename_vbah As String = "VBAHelpersTests"

Public Sub VBAHelpers_Export()
    'Exports all modules to the current directory (for source control) and increases the version number in the VBA Helpers module.

    Const versionstring_vbah As String = "'# Version "
    Const copyrightsearch_vbah As String = "'# Copyright "
    Const copyrightstring_vbah As String = "'# Copyright (c) 2012-{0} Christian Specht"
    Dim exportfile_vbah As String
    
    'export VBA Helpers
    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelpersfilename_vbah)
    Application.SaveAsText acModule, vbahelpersmodulename_vbah, exportfile_vbah

    'set version number
    Dim lines1_vbah() As String
    Dim lines2_vbah() As String
    Dim i_vbah As Long

    lines1_vbah = File_ReadAllLines(exportfile_vbah)
    ReDim lines2_vbah(UBound(lines1_vbah))
    For i_vbah = 0 To UBound(lines1_vbah)
        If String_StartsWith(lines1_vbah(i_vbah), versionstring_vbah) Then
            lines2_vbah(i_vbah) = versionstring_vbah & format(Now, "yyyymmdd.hhmmss")
        ElseIf String_StartsWith(lines1_vbah(i_vbah), copyrightsearch_vbah) Then
            lines2_vbah(i_vbah) = String_Format(copyrightstring_vbah, Year(Date))
        Else
            lines2_vbah(i_vbah) = lines1_vbah(i_vbah)
        End If
    Next

    File_WriteAllLines exportfile_vbah, lines2_vbah

    'export tests
    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelperstestfilename_vbah)
    Application.SaveAsText acModule, vbahelperstestmodulename_vbah, exportfile_vbah

    'export dev functions
    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelpersdevfilename_vbah)
    Application.SaveAsText acModule, vbahelpersdevmodulename_vbah, exportfile_vbah

End Sub

Public Sub VBAHelpers_Import()
    'Imports all VBA Helpers modules from the current directory.

    Dim exportfile_vbah As String
    Dim message_vbah As String

    'import tests
    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelperstestfilename_vbah)
    If Not File_Exists(exportfile_vbah) Then
        message_vbah = String_Format("Couldn't find test class:{0}{1}", vbCrLf, exportfile_vbah)
        MsgBox message_vbah, vbCritical
        Exit Sub
    End If
    
    Application.LoadFromText acModule, vbahelperstestmodulename_vbah, exportfile_vbah


    'import dev functions
    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelpersdevfilename_vbah)
    If Not File_Exists(exportfile_vbah) Then
        message_vbah = String_Format("Couldn't find dev functions:{0}{1}", vbCrLf, exportfile_vbah)
        MsgBox message_vbah, vbCritical
        Exit Sub
    End If
    
    Application.LoadFromText acModule, vbahelpersdevmodulename_vbah, exportfile_vbah
    

    'import VBA Helpers (use update function from actual VBA Helpers module)
    VBAHelpers_Update

End Sub