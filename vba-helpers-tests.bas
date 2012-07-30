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
    ' Arrange
    Const Expected As Boolean = True
    Dim Actual As Boolean
    ' Act
    Actual = String_EndsWith("abc", "bc")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_EndsWith_EndDoesNotMatchSecondString_ReturnsFalse()
    ' Arrange
    Const Expected As Boolean = False
    Dim Actual As Boolean
    ' Act
    Actual = String_EndsWith("abc", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_Format_StringParameter_IsInserted()
    ' Arrange
    Const Expected As String = "test x"
    Dim Actual As String
    ' Act
    Actual = String_Format("test {0}", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_Format_NumericParameter_IsInserted()
    ' Arrange
    Const Expected As String = "test 1"
    Dim Actual As String
    ' Act
    Actual = String_Format("test {0}", 1)
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_Format_MissingParameter_IsInsertedAsEmptyString()
    ' Arrange
    Const Expected As String = "test x "
    Dim Actual As String
    ' Act
    Actual = String_Format("test {0} {1}", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_Format_MissingPlaceholderAndSuppliedParameter_ParameterIsIgnored()
    ' Arrange
    Const Expected As String = "test "
    Dim Actual As String
    ' Act
    Actual = String_Format("test {0}")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_StartsWith_BeginningMatchesSecondString_ReturnsTrue()
    ' Arrange
    Const Expected As Boolean = True
    Dim Actual As Boolean
    ' Act
    Actual = String_StartsWith("abc", "ab")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Public Sub String_StartsWith_BeginningDoesNotMatchSecondString_ReturnsFalse()
    ' Arrange
    Const Expected As Boolean = False
    Dim Actual As Boolean
    ' Act
    Actual = String_StartsWith("abc", "x")
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub