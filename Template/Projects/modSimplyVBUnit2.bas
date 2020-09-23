Attribute VB_Name = "modSimplyVBUnit2"
'
' modSimplyVBUnit2 V2.0
'
Option Explicit

' This provides a simple testing harness that outputs the
' results of the test run in the debug window.
'
Private Suite As New TestSuite



Private Sub AddTests()
    ' Add any test objects to the suite here.
    '
    ' .Add <TestObject>
    '
    With Suite
        
    
    End With
End Sub
Private Sub HandleResult(ByVal Result As ITestResult)
    ' Handle the result if necessary here.
    '
    With Result
    
    
    End With
End Sub



Private Sub Main()
    Call AddTests
    Dim Result As ITestResult
    Set Result = Suite.Run(New DebugWindowListener)
    Call HandleResult(Result)
End Sub
