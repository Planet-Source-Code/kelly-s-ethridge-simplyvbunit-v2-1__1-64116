VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugWindowListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DebugWindowListener
'
' This is a simple listener used for the Debug window.
'
Option Explicit
Implements ITestListener


Private Function Format(ByVal Text As String, ParamArray Args() As Variant) As String
    Dim i As Long
    For i = 0 To UBound(Args)
        Text = Replace$(Text, "{" & i & "}", Args(i))
    Next i
    
    Format = Text
End Function

Private Sub OutputResult(ByVal Result As ITestResult)
    If Result.IsFailure Then
        Debug.Print Format("Failure In  : {0}", Result.TestName)
        Debug.Print Format("Test        : {0}", Result.Test.FullName)
        Debug.Print Format("Description : {0}", Result.Description)
        Debug.Print Format("Message     : {0}", Result.Message)
        Debug.Print
    End If
End Sub

Private Sub OutputContextFailure(ByVal Reason As String, ByVal Result As ContextResult)
    Debug.Print Format("{0} In  : {1}", Reason, Result.Name)
    Debug.Print Format("Test        : {0}", Result.Test.FullName)
    Debug.Print Format("Cause       : {0}", Result.Cause)
    Debug.Print Format("Description : {0}", Result.Description)
    Debug.Print Format("Message     : {0}", Result.Message)
    Debug.Print
End Sub

Private Sub OutputContextResult(ByVal Result As ContextResult)
    If Result.IsIgnored Then
        Call OutputContextFailure("Ignored", Result)
    ElseIf Result.IsFailure Then
        Call OutputContextFailure("Failure", Result)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestListener Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)
    Debug.Print "Message --->: " & Message
    Debug.Print
End Sub

Private Sub ITestListener_RunCompleted(ByVal Result As SimplyVBUnitLib.ITestResult)
    Debug.Print String$(80, "-")
    Debug.Print "Total Tests Run: " & Result.TestsRun
    Debug.Print "Total Failures : " & Result.FailureCount
    Debug.Print "Total Time     : " & Result.Time & " ms"
End Sub

Private Sub ITestListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    Debug.Print "Starting run in: " & Test.Name
    Debug.Print "Tests to be run: " & Test.TotalTestCount
    Debug.Print String$(80, "-")
End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If HasContext Then
        Call OutputContextResult(Result)
    End If
End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If HasContext Then
        Call OutputContextResult(Result)
    End If
End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As SimplyVBUnitLib.TestCaseResult)
    
End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As SimplyVBUnitLib.TestCase)

End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    Call OutputResult(Result)
End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If HasContext Then
        Call OutputContextResult(Result)
    End If
End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As SimplyVBUnitLib.ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If HasContext Then
        Call OutputContextResult(Result)
    End If
End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As SimplyVBUnitLib.ITest, ByVal HasSetup As Boolean)

End Sub

Private Sub ITestListener_TestStarted(ByVal Test As SimplyVBUnitLib.TestMethod)

End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Result As SimplyVBUnitLib.TestSuiteResult)

End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suite As SimplyVBUnitLib.TestSuite)

End Sub
