'##########################################################################################################################################
'#
'# VBA Helpers
'# A collection of useful VBA functions
'#
'# Version 20120725.004239
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
    'Exports the VBA Helpers module to the current directory and increases the version number.
    
    Const versionstring As String = "'# Version "
    Dim exportfile As String
    
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
    
End Sub

Public Sub VBAHelpers_Import()
    'Imports a new version of the VBA Helpers module from the current directory.

    Dim exportfile As String

    exportfile = Path_Combine(Path_GetCurrentDirectory, vbahelpersfilename)
    
    Application.LoadFromText acModule, vbahelpersmodulename, exportfile

End Sub

'##########################################################################################################################################

Public Function File_ReadAllLines(ByVal path As String) As String()
    'Reads a text file and returns a string array, each array item containing a line from the file.
    
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
    'Reads a text file and returns the content in a string variable.
    
    Dim contents() As String
    
    contents = File_ReadAllLines(path)
    
    If UBound(contents) > 0 Then
        File_ReadAllText = (Join(contents, environmentnewline))
    End If

End Function

Public Sub File_WriteAllLines(ByVal path As String, contents() As String)
    'Writes the content of a string array into a text file, each array item into a new line.

    File_WriteAllText path, Join(contents, environmentnewline)

End Sub

Public Sub File_WriteAllText(ByVal path As String, ByVal contents As String)
    'Writes the content of a string variable into a text file.
    
    Dim i As Integer
    
    i = FreeFile
    
    Close #i
    
    Open path For Output As #i
    Print #i, contents
    Close #i

End Sub

Public Function Path_Combine(ParamArray paths() As Variant) As String
    'Combines several strings into a path and takes care of directory separators, i.e. `path_combine("c:\","\foo","bar")` will return `c:\foo\bar`
    
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

Public Function Path_GetCurrentDirectory() As String
    'Returns the directory of the current Access database.
    
    Path_GetCurrentDirectory = Path_GetDirectoryName(CurrentDb.Name)
    
End Function

Public Function Path_GetDirectoryName(ByVal path As String) As String
    'Receives a complete path, returns only the directory.
    
    Dim i As Long
    
    If Len(path) > 3 Then
    
        i = InStrRev(path, directoryseparatorchar)
    
        If i > 3 Then
            Path_GetDirectoryName = Left(path, i - 1)
        End If

    End If
    
End Function

Public Function Path_GetFileName(ByVal path As String) As String
    'Receives a complete path, returns only the file name.
    
    Dim i As Long
    
    i = InStrRev(path, directoryseparatorchar)
    
    If i < Len(path) Then
        Path_GetFileName = Mid(path, i + 1)
    End If

End Function

Public Function Path_GetFileNameWithoutExtension(ByVal path As String) As String
    'Receives a complete path, returns only the file name without extension.
    
    Dim filename As String
    Dim i As Long
    
    filename = Path_GetFileName(path)
    
    i = InStrRev(filename, ".")
    
    If i > 0 Then
        Path_GetFileNameWithoutExtension = Left(filename, i - 1)
    End If
    
End Function

Public Function String_EndsWith(ByVal main As String, ByVal value As String) As Boolean
    'Returns `True` if the second parameter matches the end of the first parameter.
    
    String_EndsWith = (Right(main, Len(value)) = value)
    
End Function

Public Function String_StartsWith(ByVal main As String, ByVal value As String) As Boolean
    'Returns `True` if the second parameter matches the beginning of the first parameter.
    
    String_StartsWith = (Left(main, Len(value)) = value)
    
End Function
