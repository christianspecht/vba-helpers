Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------


Public Sub Directory_Exists_EmptyString_ReturnsFalse()

    Assert.IsFalse Directory_Exists("")

End Sub

Public Sub Directory_Exists_ExistingDir_ReturnsTrue()

    Assert.IsTrue Directory_Exists(Path_GetCurrentDirectory)

End Sub

Public Sub Directory_Exists_NonExistingDir_ReturnsFalse()

    Assert.IsFalse Directory_Exists(Path_Combine(Path_GetCurrentDirectory, "doesnt.exist"))

End Sub

Public Sub File_Exists_EmptyString_ReturnsFalse()

    Assert.IsFalse File_Exists("")

End Sub

Public Sub File_Exists_ExistingFile_ReturnsTrue()

    Assert.IsTrue File_Exists(Path_Combine(Path_GetCurrentDirectory, "readme.md"))

End Sub

Public Sub File_Exists_NonExistingFile_ReturnsFalse()

    Assert.IsFalse File_Exists(Path_Combine(Path_GetCurrentDirectory, "doesnt.exist"))

End Sub

Public Sub Path_Combine_SeparatorAtTheEnd_IsTruncated()

    Assert.AreEqualStrings Path_Combine("foo\"), "foo"
    
End Sub

Public Sub Path_Combine_TwoPathsWithoutSeparator_ResultHasOneSeparator()

    Assert.AreEqualStrings Path_Combine("foo", "bar"), "foo\bar"
    
End Sub

Public Sub Path_Combine_TwoPathsWithOneSeparator_ResultHasOneSeparator()

    Assert.AreEqualStrings Path_Combine("foo\", "bar"), "foo\bar"
    
End Sub

Public Sub Path_Combine_TwoPathsWithTwoSeparators_ResultHasOneSeparator()

    Assert.AreEqualStrings Path_Combine("foo\", "\bar"), "foo\bar"
    
End Sub

Public Sub Path_GetDirectoryName_DirectoryWithFilename_ReturnsDirectoryOnly()

    Assert.AreEqualStrings Path_GetDirectoryName("foo\bar.txt"), "foo"

End Sub

Public Sub Path_GetDirectoryName_RootDirectoryOnly_ReturnsEmptyString()

    Assert.AreEqualStrings Path_GetDirectoryName("c:\"), ""

End Sub

Public Sub Path_GetExtension_DirectoryWithFilename_ReturnsExtensionWithDot()

    Assert.AreEqualStrings Path_GetExtension("foo\bar.txt"), ".txt"
 
End Sub

Public Sub Path_GetExtension_DirectoryWithFilenameWithoutExtension_ReturnsEmptyString()

    Assert.AreEqualStrings Path_GetExtension("foo\bar"), ""
 
End Sub

Public Sub Path_GetExtension_DirectoryWithDotAndFilenameWithoutExtension_ReturnsEmptyString()

    Assert.AreEqualStrings Path_GetExtension("fo.o\bar"), ""
 
End Sub

Public Sub Path_GetFileName_DirectoryWithFilename_ReturnsFilenameOnly()

    Assert.AreEqualStrings Path_GetFileName("foo\bar.txt"), "bar.txt"
 
End Sub

Public Sub Path_GetFileName_DirectoryWithFilenameWithoutExtension_ReturnsFilenameWithoutExtensionOnly()

    Assert.AreEqualStrings Path_GetFileName("foo\bar"), "bar"
 
End Sub

Public Sub Path_GetFileName_PathWithoutSeparator_ReturnsEmptyString()

    Assert.AreEqualStrings Path_GetFileName("foo"), ""
 
End Sub

Public Sub Path_GetFileNameWithoutExtension_FilenameWithExtension_ReturnsFilenameOnly()

    Assert.AreEqualStrings Path_GetFileNameWithoutExtension("foo\bar.ext"), "bar"
    
End Sub

Public Sub Path_GetFileNameWithoutExtension_FilenameWithoutExtension_ReturnsFilename()

    Assert.AreEqualStrings Path_GetFileNameWithoutExtension("foo\bar"), "bar"
    
End Sub

Public Sub Path_GetFileNameWithoutExtension_DirectoryWithDotAndFilenameWithoutExtension_ReturnsFileName()

    Assert.AreEqualStrings Path_GetFileNameWithoutExtension("fo.o\bar"), "bar"
    
End Sub

Public Sub String_Contains_ContainsString_ReturnsTrue()

    Assert.IsTrue String_Contains("abc", "ab")
    
End Sub

Public Sub String_Contains_DoesNotContainString_ReturnsFalse()

    Assert.IsFalse String_Contains("abc", "ac")
    
End Sub

Public Sub String_EndsWith_EndMatchesSecondString_ReturnsTrue()

    Assert.IsTrue String_EndsWith("abc", "bc")
    
End Sub

Public Sub String_EndsWith_EndDoesNotMatchSecondString_ReturnsFalse()

    Assert.IsFalse String_EndsWith("abc", "x")

End Sub

Public Sub String_Format_StringParameter_IsInserted()

    Assert.AreEqualStrings String_Format("test {0}", "x"), "test x"
    
End Sub

Public Sub String_Format_NumericParameter_IsInserted()

    Assert.AreEqualStrings String_Format("test {0}", 1), "test 1"
    
End Sub

Public Sub String_Format_MissingParameter_IsInsertedAsEmptyString()

    Assert.AreEqualStrings String_Format("test {0} {1}", "x"), "test x "
    
End Sub

Public Sub String_Format_MissingPlaceholderAndSuppliedParameter_ParameterIsIgnored()

    Assert.AreEqualStrings String_Format("test {0}"), "test "
    
End Sub

Public Sub String_IsNullOrEmpty_Null_ReturnsTrue()

    Assert.IsTrue String_IsNullOrEmpty(Null)

End Sub

Public Sub String_IsNullOrEmpty_EmptyString_ReturnsTrue()

    Assert.IsTrue String_IsNullOrEmpty("")

End Sub

Public Sub String_IsNullOrEmpty_NullString_ReturnsTrue()

    Assert.IsTrue String_IsNullOrEmpty(vbNullString)

End Sub

Public Sub String_IsNullOrEmpty_NonEmptyString_ReturnsFalse()

    Assert.IsFalse String_IsNullOrEmpty("foo")

End Sub

Public Sub String_IsNullOrEmpty_Blank_ReturnsFalse()

    Assert.IsFalse String_IsNullOrEmpty(" ")

End Sub

Public Sub String_IsNullOrWhiteSpace_Null_ReturnsTrue()

    Assert.IsTrue String_IsNullOrWhiteSpace(Null)

End Sub

Public Sub String_IsNullOrWhiteSpace_EmptyString_ReturnsTrue()

    Assert.IsTrue String_IsNullOrWhiteSpace("")

End Sub

Public Sub String_IsNullOrWhiteSpace_NullString_ReturnsTrue()

    Assert.IsTrue String_IsNullOrWhiteSpace(vbNullString)

End Sub

Public Sub String_IsNullOrWhiteSpace_NonEmptyString_ReturnsFalse()

    Assert.IsFalse String_IsNullOrWhiteSpace("foo")

End Sub

Public Sub String_IsNullOrWhiteSpace_Blank_ReturnsTrue()

    Assert.IsTrue String_IsNullOrWhiteSpace(" ")

End Sub

Public Sub String_PadLeft_WithPaddingChar_IsPaddedCorrectly()

    Assert.AreEqualStrings String_PadLeft("foo", 5, "a"), "aafoo"

End Sub

Public Sub String_PadLeft_WithPaddingString_IsPaddedWithFirstChar()

    Assert.AreEqualStrings String_PadLeft("foo", 5, "ab"), "aafoo"

End Sub

Public Sub String_PadLeft_WithoutPaddingChar_IsPaddedCorrectly()

    Assert.AreEqualStrings String_PadLeft("foo", 5), "  foo"

End Sub

Public Sub String_PadRight_WithPaddingChar_IsPaddedCorrectly()

    Assert.AreEqualStrings String_PadRight("foo", 5, "a"), "fooaa"

End Sub

Public Sub String_PadRight_WithPaddingString_IsPaddedWithFirstChar()

    Assert.AreEqualStrings String_PadRight("foo", 5, "ab"), "fooaa"

End Sub

Public Sub String_PadRight_WithoutPaddingChar_IsPaddedCorrectly()

    Assert.AreEqualStrings String_PadRight("foo", 5), "foo  "

End Sub

Public Sub String_StartsWith_BeginningMatchesSecondString_ReturnsTrue()

    Assert.IsTrue String_StartsWith("abc", "ab")
    
End Sub

Public Sub String_StartsWith_BeginningDoesNotMatchSecondString_ReturnsFalse()

    Assert.IsFalse String_StartsWith("abc", "x")

End Sub