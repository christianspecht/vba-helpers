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


Public Sub String_EndsWith_EndMatchesSecondString_ReturnsTrue()

    Assert.That String_EndsWith("abc", "bc"), Iz.EqualTo(True)
    
End Sub

Public Sub String_EndsWith_EndDoesNotMatchSecondString_ReturnsFalse()

    Assert.That String_EndsWith("abc", "x"), Iz.EqualTo(False)

End Sub

Public Sub String_Format_StringParameter_IsInserted()

    Assert.That String_Format("test {0}", "x"), Iz.EqualTo("test x")
    
End Sub

Public Sub String_Format_NumericParameter_IsInserted()

    Assert.That String_Format("test {0}", 1), Iz.EqualTo("test 1")
    
End Sub

Public Sub String_Format_MissingParameter_IsInsertedAsEmptyString()

    Assert.That String_Format("test {0} {1}", "x"), Iz.EqualTo("test x ")
    
End Sub

Public Sub String_Format_MissingPlaceholderAndSuppliedParameter_ParameterIsIgnored()

    Assert.That String_Format("test {0}"), Iz.EqualTo("test ")
    
End Sub

Public Sub String_StartsWith_BeginningMatchesSecondString_ReturnsTrue()

    Assert.That String_StartsWith("abc", "ab"), Iz.EqualTo(True)
    
End Sub

Public Sub String_StartsWith_BeginningDoesNotMatchSecondString_ReturnsFalse()

    Assert.That String_StartsWith("abc", "x"), Iz.EqualTo(False)

End Sub