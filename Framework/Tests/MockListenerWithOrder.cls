VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockListenerWithOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MockListenerWithOrder
'
Option Explicit
Implements ITestListener

Public TestFixtureSetupStarted As Long
Public TestFixtureSetupCompleted As Long
Public SetupStarted As Long
Public SetupCompleted As Long
Public TestStarted As Long
Public TestCompleted As Long
Public TeardownStarted As Long
Public TeardownCompleted As Long
Public TestFixtureTeardownStarted As Long
Public TestFixtureTeardownCompleted As Long
Public Hear  As Long
Public RunCompleted As Long
Public RunStarted As Long
Public TestCaseStarted As Long
Public TestCaseCompleted As Long
Public TestSuiteStarted As Long
Public TestSuiteCompleted As Long

Private seq As Long



Private Function NextSeq() As Long
    seq = seq + 1
    NextSeq = seq
End Function

Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)
    Hear = NextSeq
End Sub

Private Sub ITestListener_RunCompleted(ByVal Result As SimplyVBUnitLib.ITestResult)
    RunCompleted = NextSeq
End Sub

Private Sub ITestListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    RunStarted = NextSeq
End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    SetupCompleted = NextSeq
End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)
    SetupStarted = NextSeq
End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    TeardownCompleted = NextSeq
End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)
    TeardownStarted = NextSeq
End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As SimplyVBUnitLib.TestCaseResult)
    TestCaseCompleted = NextSeq
End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As SimplyVBUnitLib.TestCase)
    TestCaseStarted = NextSeq
End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    TestCompleted = NextSeq
End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    TestFixtureSetupCompleted = NextSeq
End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As SimplyVBUnitLib.ITest, ByVal HasContext As Boolean)
    TestFixtureSetupStarted = NextSeq
End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    TestFixtureTeardownCompleted = NextSeq
End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As SimplyVBUnitLib.ITest, ByVal HasSetup As Boolean)
    TestFixtureTeardownStarted = NextSeq
End Sub

Private Sub ITestListener_TestStarted(ByVal Test As SimplyVBUnitLib.TestMethod)
    TestStarted = NextSeq
End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Result As SimplyVBUnitLib.TestSuiteResult)
    TestSuiteCompleted = NextSeq
End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suite As SimplyVBUnitLib.TestSuite)
    TestSuiteStarted = NextSeq
End Sub
