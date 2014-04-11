'##########################################################################################################################################
'#
'# VBA Helpers
'# A collection of useful VBA functions
'#
'# Version 20140411.200133
'# (the version number is just the current date/time)
'#
'# Copyright (c) 2012-2014 Christian Specht
'#
'# Visit the project site for documentation and more information:
'# http://christianspecht.de/vba-helpers/
'#
'# VBA Helpers is licensed under the MIT License:
'# http://christianspecht.de/vba-helpers/#license
'#
'##########################################################################################################################################

Option Compare Database
Option Explicit

Public Const vbahelpersfilename_vbah As String = "vba-helpers.bas"
Public Const vbahelpersmodulename_vbah As String = "VBAHelpers"

Const directoryseparatorchar_vbah As String = "\"
Const environmentnewline_vbah As String = vbCrLf

'return value for Environment_AccessVersion()
Public Enum accessversion_vbah
    Access1995 = 7
    Access1997 = 8
    Access2000 = 9
    Access2002 = 10
    Access2003 = 11
    Access2007 = 12
    Access2010 = 14
    Access2013 = 15
End Enum

'API call to ShellExecute, needed for Process_Start()
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long


'##########################################################################################################################################


Public Function Directory_Exists(ByVal path_vbah As String) As Boolean
    'Returns True if the specified directory exists.
    
    If path_vbah > "" Then
        Directory_Exists = (Dir(path_vbah, vbDirectory) > "")
    End If
    
End Function

Public Function Environment_AccessVersion() As accessversion_vbah
    'Returns an Enum which contains the version of the current msaccess.exe.

    Environment_AccessVersion = Val(SysCmd(acSysCmdAccessVer))

End Function

Public Function Environment_MachineName() As String
    'Returns the name of the local computer.
    
    Environment_MachineName = Environ("computername")
    
End Function

Public Sub File_Delete(ByVal path_vbah As String)
    'Deletes a file. If the file does not exist, nothing happens.

    If File_Exists(path_vbah) Then
        Kill path_vbah
    End If

End Sub

Public Function File_Exists(ByVal path_vbah As String) As Boolean
    'Returns `True` if the specified file exists.
    
    If path_vbah > "" Then
        File_Exists = (Dir(path_vbah) > "")
    End If
    
End Function

Public Function File_ReadAllLines(ByVal path_vbah As String) As String()
    'Reads a text file and returns a string array, each array item containing a line from the file.
    
    Dim i_vbah As Integer
    Dim tmp_vbah As String
    Dim filelines_vbah As Long
    Dim arraylines_vbah As Long
    Dim retval_vbah() As String
    
    i_vbah = FreeFile
    Close #i_vbah
    
    Open path_vbah For Input As #i_vbah
    
    filelines_vbah = 0
    arraylines_vbah = 0
    
    Do While Not EOF(i_vbah)
        
        If arraylines_vbah <= filelines_vbah Then
            arraylines_vbah = arraylines_vbah + 100
            ReDim Preserve retval_vbah(arraylines_vbah - 1)
        End If
        
        Line Input #i_vbah, tmp_vbah
        retval_vbah(filelines_vbah) = tmp_vbah
        
        filelines_vbah = filelines_vbah + 1
        
    Loop
    
    ReDim Preserve retval_vbah(filelines_vbah - 1)
    
    Close #i_vbah
    
    File_ReadAllLines = retval_vbah
    
End Function

Public Function File_ReadAllText(ByVal path_vbah As String) As String
    'Reads a text file and returns the content in a string variable.
    
    Dim contents_vbah() As String
    
    contents_vbah = File_ReadAllLines(path_vbah)
    
    If UBound(contents_vbah) > 0 Then
        File_ReadAllText = (Join(contents_vbah, environmentnewline_vbah))
    End If

End Function

Public Sub File_WriteAllLines(ByVal path_vbah As String, contents_vbah() As String)
    'Writes the content of a string array into a text file, each array item into a new line.

    File_WriteAllText path_vbah, Join(contents_vbah, environmentnewline_vbah)

End Sub

Public Sub File_WriteAllText(ByVal path_vbah As String, ByVal contents_vbah As String)
    'Writes the content of a string variable into a text file.
    
    Dim i_vbah As Integer
    
    i_vbah = FreeFile
    
    Close #i_vbah
    
    Open path_vbah For Output As #i_vbah
    Print #i_vbah, contents_vbah
    Close #i_vbah

End Sub

Public Function InputBox_PressedCancel(ByRef inputboxresult_vbah) As Boolean
    'Receives the return value of an InputBox, returns True when the input was canceled.
    'Normally you can't distinguish whether you cancelled the input or submitted an empty string - the InputBox returns an empty string in both cases.
    'Example: InputBox_PressedCancel(InputBox("foo")) returns True when you press Cancel, and False when you press OK without entering a value.

    InputBox_PressedCancel = (StrPtr(inputboxresult_vbah) = 0)

End Function

Public Function Path_Combine(ParamArray paths_vbah() As Variant) As String
    'Combines several strings into a path and takes care of directory separators.
    'Example: `path_combine("c:\","\foo","bar")` will return `c:\foo\bar`

    Dim path_vbah As Variant
    Dim retval_vbah As String
    
    For Each path_vbah In paths_vbah
    
        If String_StartsWith(path_vbah, directoryseparatorchar_vbah) Then
            path_vbah = Mid(path_vbah, Len(directoryseparatorchar_vbah) + 1)
        End If
    
        If String_EndsWith(path_vbah, directoryseparatorchar_vbah) Then
            path_vbah = Left(path_vbah, Len(path_vbah) - Len(directoryseparatorchar_vbah))
        End If
    
        retval_vbah = retval_vbah & path_vbah & directoryseparatorchar_vbah
    
    Next
    
    If String_EndsWith(retval_vbah, directoryseparatorchar_vbah) Then
        retval_vbah = Left(retval_vbah, Len(retval_vbah) - Len(directoryseparatorchar_vbah))
    End If
    
    Path_Combine = retval_vbah

End Function

Public Function Path_GetCurrentDirectory() As String
    'Returns the directory of the current Access database.
    
    Path_GetCurrentDirectory = Path_GetDirectoryName(CurrentDb.Name)
    
End Function

Public Function Path_GetDirectoryName(ByVal path_vbah As String) As String
    'Receives a complete path, returns only the directory.
    
    Dim i_vbah As Long
    
    If Len(path_vbah) > 3 Then
    
        i_vbah = InStrRev(path_vbah, directoryseparatorchar_vbah)
    
        If i_vbah > 3 Then
            Path_GetDirectoryName = Left(path_vbah, i_vbah - 1)
        End If

    End If
    
End Function

Public Function Path_GetExtension(ByVal path_vbah As String) As String
    'Receives a complete path, returns only the extension.
    
    Dim filename_vbah As String
    Dim i_vbah As Long
    
    filename_vbah = Path_GetFileName(path_vbah)
    
    i_vbah = InStrRev(filename_vbah, ".")
    
    If i_vbah > 0 Then
        Path_GetExtension = Mid(filename_vbah, i_vbah)
    End If
    
End Function

Public Function Path_GetFileName(ByVal path_vbah As String) As String
    'Receives a complete path, returns only the file name.
    
    Dim i_vbah As Long
    
    i_vbah = InStrRev(path_vbah, directoryseparatorchar_vbah)
    
    If i_vbah > 0 And i_vbah < Len(path_vbah) Then
        Path_GetFileName = Mid(path_vbah, i_vbah + 1)
    End If

End Function

Public Function Path_GetFileNameWithoutExtension(ByVal path_vbah As String) As String
    'Receives a complete path, returns only the file name without extension.
    
    Dim filename_vbah As String
    Dim i_vbah As Long
    
    filename_vbah = Path_GetFileName(path_vbah)
    
    i_vbah = InStrRev(filename_vbah, ".")
    
    If i_vbah = 0 Then
        Path_GetFileNameWithoutExtension = filename_vbah
    ElseIf i_vbah > 0 Then
        Path_GetFileNameWithoutExtension = Left(filename_vbah, i_vbah - 1)
    End If
    
End Function

Public Function Path_GetTempPath() As String
    'Returns the current user's temp folder.
    
    Path_GetTempPath = Environ("temp")
    
End Function

Public Sub Process_Start(ByVal path_vbah, Optional ByVal parameters_vbah As String, Optional ByVal hidewindow_vbah As Boolean)
    'Executes a file. If the file itself is not an application, it will be started with the default application (as if you double-clicked it in Windows Explorer).
    'Use the optional parameters to supply command-line arguments to the executed file, and to open the file hidden (without a visible window - useful for executing command-line tools)
    
    If File_Exists(path_vbah) Then
        ShellExecute 0, "open", path_vbah, parameters_vbah, "", IIf(hidewindow_vbah, 0, 1)
    End If
    
End Sub

Public Function String_Contains(ByVal main_vbah As String, ByVal value_vbah As String) As Boolean
    'Returns `True` if the second parameter occurs within the first parameter.
    'Example: `String_Contains("abc", "ab")` will return `True`
    
    String_Contains = (InStr(1, main_vbah, value_vbah) > 0)
    
End Function

Public Function String_EndsWith(ByVal main_vbah As String, ByVal value_vbah As String) As Boolean
    'Returns `True` if the second parameter matches the end of the first parameter.
    'Example: `String_EndsWith("abc", "bc")` will return `True`
    
    String_EndsWith = (Right(main_vbah, Len(value_vbah)) = value_vbah)
    
End Function

Public Function String_Format(ByVal format_vbah As String, ParamArray args_vbah() As Variant)
    'Replaces numbered placeholders (`{0}`, `{1}`, ...) in the first parameter by the corresponding value from the additional parameter list.
    'Example: `String_Format("Hello {0}", "world")` will return `Hello world`

    Dim numberofargs_vbah As Integer
    Dim i_vbah As Integer
    
    numberofargs_vbah = UBound(args_vbah)
    
    For i_vbah = 0 To 100
    
        If i_vbah <= numberofargs_vbah Then
            format_vbah = Replace(format_vbah, "{" & i_vbah & "}", args_vbah(i_vbah))
        Else
            format_vbah = Replace(format_vbah, "{" & i_vbah & "}", "")
        End If
    
    Next
    
    String_Format = format_vbah

End Function

Public Function String_IsNullOrEmpty(ByVal input_vbah As Variant) As Boolean
    'Returns True when the input is either Null or an empty string.
    '(note: a VBA string can't be Null, but the function is called `String_` anyway to keep the naming consistent)

    String_IsNullOrEmpty = (Nz(input_vbah) = "")

End Function

Public Function String_IsNullOrWhiteSpace(ByVal input_vbah As Variant) As Boolean
    'Returns True when the input is either Null, an empty string or consists of whitespace characters (blanks) only.
    '(note: a VBA string can't be Null, but the function is called `String_` anyway to keep the naming consistent)
    
    Dim retval_vbah As Boolean
    
    retval_vbah = String_IsNullOrEmpty(input_vbah)
    
    If Not retval_vbah Then
        If Trim(input_vbah) = "" Then
            retval_vbah = True
        End If
    End If

    String_IsNullOrWhiteSpace = retval_vbah
    
End Function

Public Function String_PadLeft(ByVal inputstring_vbah, ByVal totalwidth_vbah, Optional ByVal paddingchar_vbah = " ")
    'Right-aligns the first string parameter by padding it on the left with the second string parameter, up to the total specified width.
    'Example: `String_PadLeft("foo",5,"a")` will return `aafoo`
    
    String_PadLeft = Right(String(totalwidth_vbah, Left(paddingchar_vbah, 1)) & inputstring_vbah, totalwidth_vbah)
    
End Function

Public Function String_PadRight(ByVal inputstring_vbah, ByVal totalwidth_vbah, Optional ByVal paddingchar_vbah = " ")
    'Left-aligns the first string parameter by padding it on the right with the second string parameter, up to the total specified width.
    'Example: `String_PadRight("foo",5,"a")` will return `fooaa`
    
    String_PadRight = Left(inputstring_vbah & String(totalwidth_vbah, Left(paddingchar_vbah, 1)), totalwidth_vbah)
    
End Function

Public Function String_StartsWith(ByVal main_vbah As String, ByVal value_vbah As String) As Boolean
    'Returns `True` if the second parameter matches the beginning of the first parameter.
    'Example: `String_StartsWith("abc", "ab")` will return `True`
    
    String_StartsWith = (Left(main_vbah, Len(value_vbah)) = value_vbah)
    
End Function

Public Function VBAHelpers_Update()
    'Updates VBA Helpers to newer version by importing a downloaded file (file must be in same folder as current Access database)
    
    Dim exportfile_vbah As String
    Dim message_vbah As String

    exportfile_vbah = Path_Combine(Path_GetCurrentDirectory, vbahelpersfilename_vbah)

    If Not File_Exists(exportfile_vbah) Then
        message_vbah = String_Format("Couldn't find VBA Helpers file in current directory:{0}{1}{0}{0}VBA Helpers update failed!", vbCrLf, exportfile_vbah)
        MsgBox message_vbah, vbCritical
        Exit Function
    End If

    Application.LoadFromText acModule, vbahelpersmodulename_vbah, exportfile_vbah

End Function
