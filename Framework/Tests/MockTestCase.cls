VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockTestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MockTestCase
'
Option Explicit

Public TestOneCalled As Boolean
Public TestTwoCalled As Boolean
Public TestThreeCalled As Boolean
Public TestFourCalled As Boolean


Public Sub TestOne()
    TestOneCalled = True
End Sub

Public Sub TestTwo()
    TestTwoCalled = True
End Sub

Public Sub TestThree()
    TestThreeCalled = True
    Assert.Fail "This test should fail."
End Sub

Public Sub TestFour()
    TestFourCalled = True
    Assert.AreEqual 1, 0, "This test should fail."
End Sub
