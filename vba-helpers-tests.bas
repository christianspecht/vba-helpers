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


Public Sub String_StartsWith_BeginningMatchesSecondString_ReturnsTrue()
    ' Arrange
    Const Expected As Variant = True
    Dim Actual As Variant
    ' Act
    Actual = String_StartsWith("abc", "ab")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_StartsWith_BeginningDoesNotMatchSecondString_ReturnsFalse()
    ' Arrange
    Const Expected As Variant = False
    Dim Actual As Variant
    ' Act
    Actual = String_StartsWith("abc", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_EndsWith_EndMatchesSecondString_ReturnsTrue()
    ' Arrange
    Const Expected As Variant = True
    Dim Actual As Variant
    ' Act
    Actual = String_EndsWith("abc", "bc")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_EndsWith_EndDoesNotMatchSecondString_ReturnsFalse()
    ' Arrange
    Const Expected As Variant = False
    Dim Actual As Variant
    ' Act
    Actual = String_EndsWith("abc", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub