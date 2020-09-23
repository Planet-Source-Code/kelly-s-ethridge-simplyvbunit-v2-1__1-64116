VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCore
'
' This class is used to test the minimum functionality
' required to allow SimplyVBUnitLib to test itself.
'
' An external tester is required to prevent SimplyVBUnitLib
' from just faking test results.
'
Option Explicit

''
' We run our core functionality tests in this one class.
'
Public Sub Run()
    TestAssertFailIsThrown
    'TestAssertFaileIsThrownWithMessage
    
    ' These tests can start using Assert.Fail
    MakeSureTestSuiteRunsUserTests
    MakeSureAllUserTestsAreRun
End Sub

''
' Assert.Fail is the primary way to indicate that an assertion
' has failed. This notifies SimplyVBUnitLib that something failed.
' All other assertions will rely on Assert.Fail working correctly.
'
Private Sub TestAssertFailIsThrown()
    On Error GoTo errTrap
    Assert.Fail
    On Error GoTo 0
    Err.Raise 5, , "Assert.Fail should raise an error."
    
errTrap:
    If Err.Number <> Assert.AssertCode Then
        Err.Raise 5, , "The wrong error code was raised by Assert.Fail."
    End If
End Sub

' <obsolete>
'
'''
'' This is the same test as TestAssertFailIsThrown with
'' a user specified message included.
''
'Private Sub TestAssertFaileIsThrownWithMessage()
'    Const Message As String = "Fail should be called."
'
'    On Error GoTo errTrap
'    Assert.Fail Message
'    On Error GoTo 0
'    Err.Raise 5, , "Assert.Fail should raise an error."
'
'errTrap:
'    If Err.Number <> Assert.AssertCode Then
'        Err.Raise 5, , "The wrong error code was raised by Assert.Fail."
'    End If
'    If Err.Description <> Message Then
'        Err.Raise 5, , "The wrong message was sent with the error."
'    End If
'End Sub

Private Sub MakeSureTestSuiteRunsUserTests()
    Dim Suite As New TestSuite
    Dim Tests As New TestSuiteRunsTests
    Suite.Add Tests
    Suite.Run
    Tests.Verify
End Sub

Private Sub MakeSureAllUserTestsAreRun()
    Dim Suite As New TestSuite
    Dim test1 As New TestSuiteRunsTests
    Dim test2 As New TestSuiteRunsTests
    Suite.Add test1, "test1"
    Suite.Add test2, "test2"
    Suite.Run
    test1.Verify
    test2.Verify
End Sub
