'##########################################################################################################################################
'#
'# VBA Helpers
'# A collection of useful VBA functions
'#
'# Version 20120724.005041
'# (the version number is just the current date/time)
'#
'# Copyright (c) 2012 Christian Specht
'# http://christianspecht.de/vba-helpers/
'#
'# VBA Helpers is licensed under the MIT License.
'# See https://bitbucket.org/christianspecht/vba-helpers/raw/tip/license.txt for details.
'#
'##########################################################################################################################################

Option Compare Database
Option Explicit

Const vbahelpersfilename As String = "vba-helpers.bas"
Const vbahelpersmodulename As String = "VBAHelpers"

Const directoryseparatorchar As String = "\"
Const environmentnewline As String = vbCrLf

'##########################################################################################################################################
'Helper functions for exporting/importing VBA Helpers itself (for source control)

Public Sub VBAHelpers_Export()
    'Exports the whole module to the folder of the current database and sets the version number.
    
    Const versionstring As String = "'# Version "
    Dim exportfile As String
    
    exportfile = Path_Combine(Path_GetCurrentPath, vbahelpersfilename)
    
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
    
End Sub

Public Sub VBAHelpers_Import()
    'Imports the module from the folder of the current database

    Dim exportfile As String

    exportfile = Path_Combine(Path_GetCurrentPath, vbahelpersfilename)
    
    Application.LoadFromText acModule, vbahelpersmodulename, exportfile

End Sub

'##########################################################################################################################################

Public Sub File_WriteAllLines(ByVal path As String, contents() As String)

    File_WriteAllText path, Join(contents, environmentnewline)

End Sub

Public Sub File_WriteAllText(ByVal path As String, ByVal contents As String)

    Dim i As Integer
    
    i = FreeFile
    
    Close #i
    
    Open path For Output As #i
    Print #i, contents
    Close #i

End Sub

Public Function File_ReadAllLines(ByVal path As String) As String()

    Dim i As Integer
    Dim tmp As String
    Dim filelines As Long
    Dim arraylines As Long
    Dim retval() As String
    
    i = FreeFile
    
    Close #i
    
    Open path For Input As #i
    
    filelines = 0
    arraylines = 0
    
    Do While Not EOF(i)
        
        If arraylines <= filelines Then
            arraylines = arraylines + 100
            ReDim Preserve retval(arraylines - 1)
        End If
        
        Line Input #i, tmp
        retval(filelines) = tmp
        
        filelines = filelines + 1
        
    Loop
    
    ReDim Preserve retval(filelines - 1)
    
    Close #i
    
    File_ReadAllLines = retval
    
End Function

Public Function File_ReadAllText(ByVal path As String) As String

    Dim contents() As String
    
    contents = File_ReadAllLines(path)
    
    If UBound(contents) > 0 Then
        File_ReadAllText = (Join(contents, environmentnewline))
    End If

End Function

Public Function Path_Combine(ParamArray paths() As Variant) As String

    Dim path As Variant
    Dim retval As String
    
    For Each path In paths
    
        If String_StartsWith(path, directoryseparatorchar) Then
            path = Mid(path, Len(directoryseparatorchar) + 1)
        End If
    
        If String_EndsWith(path, directoryseparatorchar) Then
            path = Left(path, Len(path) - Len(directoryseparatorchar))
        End If
    
        retval = retval & path & directoryseparatorchar
    
    Next
    
    If String_EndsWith(retval, directoryseparatorchar) Then
        retval = Left(retval, Len(retval) - Len(directoryseparatorchar))
    End If
    
    Path_Combine = retval

End Function

Public Function Path_GetCurrentPath() As String
    
    Path_GetCurrentPath = Path_GetDirectoryName(CurrentDb.Name)
    
End Function

Public Function Path_GetDirectoryName(ByVal path As String) As String

    Dim i As Long
    
    If Len(path) > 3 Then
    
        i = InStrRev(path, directoryseparatorchar)
    
        If i > 3 Then
            Path_GetDirectoryName = Left(path, i - 1)
        End If

    End If
    
End Function

Public Function Path_GetFileName(ByVal path As String) As String

    Dim i As Long
    
    i = InStrRev(path, directoryseparatorchar)
    
    If i < Len(path) Then
        Path_GetFileName = Mid(path, i + 1)
    End If

End Function

Public Function Path_GetFileNameWithoutExtension(ByVal path As String) As String

    Dim filename As String
    Dim i As Long
    
    filename = Path_GetFileName(path)
    
    i = InStrRev(filename, ".")
    
    If i > 0 Then
        Path_GetFileNameWithoutExtension = Left(filename, i - 1)
    End If
    
End Function

Public Function String_EndsWith(ByVal main As String, ByVal value As String) As Boolean
    String_EndsWith = (Right(main, Len(value)) = value)
End Function

Public Function String_StartsWith(ByVal main As String, ByVal value As String) As Boolean
    String_StartsWith = (Left(main, Len(value)) = value)
End Function
