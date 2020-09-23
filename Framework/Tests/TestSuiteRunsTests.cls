VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSuiteRunsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TestOneWasRun As Boolean
Public TestTwoWasRun As Boolean


Friend Sub Verify()
    If Not TestOneWasRun Then
        Assert.Fail "TestOne should have been run."
    End If
    If Not TestTwoWasRun Then
        Assert.Fail "TestTwo should have been run."
    End If
End Sub

Public Sub TestOne()
    TestOneWasRun = True
End Sub

Public Sub TestTwo()
    TestTwoWasRun = True
End Sub

Public Sub TestThree(ByVal Arg As Long)
    Assert.Fail "TestSuiteRunsTests.TestThree should not be run."
End Sub

Public Function TestFour() As Long
    Assert.Fail "TestSuiteRunsTests.TestFour should not be run."
End Function
