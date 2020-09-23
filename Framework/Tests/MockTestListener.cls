VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockTestListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MockTestListener
'
Option Explicit
Implements ITestListener

Public CompletedCount           As Long
Public StartingCount            As Long
Public TestCaseCompletedCount   As Long
Public TestCaseStartingCount    As Long
Public TestCompletedCount       As Long
Public TestStartingCount        As Long
Public TestSuiteCompletedCount  As Long
Public TestSuiteStartingCount   As Long
Public HearCount                As Long


Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)
    HearCount = HearCount + 1
End Sub

Private Sub ITestListener_RunCompleted(ByVal Results As SimplyVBUnitLib.ITestResult)
    Inc CompletedCount
End Sub

Private Sub ITestListener_RunStarted(ByVal Test As ITest)
    Inc StartingCount
End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As TestCaseResult)
    Inc TestCaseCompletedCount
End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As TestCase)
    Inc TestCaseStartingCount
End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    Inc TestCompletedCount
End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestStarted(ByVal Test As TestMethod)
    Inc TestStartingCount
End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Results As SimplyVBUnitLib.TestSuiteResult)
    Inc TestSuiteCompletedCount
End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suite As TestSuite)
    Inc TestSuiteStartingCount
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Privet Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Inc(ByRef Count As Long)
    Count = Count + 1
End Sub
