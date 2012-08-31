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