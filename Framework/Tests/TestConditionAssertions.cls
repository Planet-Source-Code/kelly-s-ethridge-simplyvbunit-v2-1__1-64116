VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestConditionAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestConditionAssertions
'
Option Explicit

Public Sub TestFailingIsTrue()
    On Error GoTo errTrap
    Assert.IsTrue False
    On Error GoTo 0
    Assert.Fail "IsTrue should have raised an error."
    
errTrap:
    If Err.Number <> Assert.AssertCode Then
        Assert.Fail "Wrong error code raised by IsTrue."
    End If
    
End Sub

Public Sub TestPassingIsTrue()
    On Error GoTo errTrap
    Assert.IsTrue True
    Exit Sub

errTrap:
    Assert.Fail "IsTrue should not have raised an error."
End Sub

'Public Sub TestFailingIsTrueWithMessage()
'    On Error GoTo errTrap
'    Assert.IsTrue False, "IsTrue should fail."
'    On Error GoTo 0
'    Assert.Fail "IsTrue should have raised an error."
'
'errTrap:
'    If Err.Number <> Assert.AssertCode Then
'        Assert.Fail "Wrong error code raised by IsTrue."
'    End If
'    If Err.Description <> "IsTrue should fail." Then
'        Assert.Fail "Wrong message returned by IsTrue."
'    End If
'End Sub

Public Sub TestFailingIsFalse()
    On Error GoTo errTrap
    Assert.IsFalse True
    On Error GoTo 0
    Assert.Fail "IsFalse should have raised an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsFalse()
    On Error GoTo errTrap
    Assert.IsFalse False
    Exit Sub

errTrap:
    Assert.Fail "IsFalse should not have raised an error."
End Sub

Public Sub TestPassingIsEmpty()
    On Error GoTo errTrap
    Assert.IsEmpty Empty
    Exit Sub

errTrap:
    Assert.Fail "IsEmpty should not raise an error."
End Sub

Public Sub TestFailingIsEmpty()
    On Error GoTo errTrap
    Assert.IsEmpty 1
    On Error GoTo 0
    Assert.Fail "IsEmpty should have raised an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsNotEmpty()
    On Error GoTo errTrap
    Assert.IsNotEmpty 1
    Exit Sub

errTrap:
    Assert.Fail "IsNotEmpty should not have raised an error."
End Sub

Public Sub TestFailingIsNotEmpty()
    On Error GoTo errTrap
    Assert.IsNotEmpty Empty
    On Error GoTo 0
    Assert.Fail "IsNotEmpty should raise an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsNull()
    On Error GoTo errTrap
    Assert.IsNull Null
    Exit Sub

errTrap:
    Assert.Fail "IsNull should not raise an error."
End Sub

Public Sub TestFailingIsNull()
    On Error GoTo errTrap
    Assert.IsNull 1
    On Error GoTo 0
    Assert.Fail "IsNull should raise an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsNotNull()
    On Error GoTo errTrap
    Assert.IsNotNull 1
    Exit Sub

errTrap:
    Assert.Fail "IsNotNull should have raised an error."
End Sub

Public Sub TestFailingIsNotNull()
    On Error GoTo errTrap
    Assert.IsNotNull Null
    On Error GoTo 0
    Assert.Fail "IsNotNull should have raised an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsNothing()
    On Error GoTo errTrap
    Assert.IsNothing Nothing
    Exit Sub

errTrap:
    Assert.Fail "IsNothing should not raise an error."
End Sub

Public Sub TestFailingIsNothing()
    On Error GoTo errTrap
    Assert.IsNothing Me
    On Error GoTo 0
    Assert.Fail "IsNothing should raise an error."

errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingIsNotNothing()
    On Error GoTo errTrap
    Assert.IsNotNothing Me
    Exit Sub

errTrap:
    Assert.Fail "IsNotNothing should not raise an error."
End Sub

Public Sub TestFailingIsNotNothing()
    On Error GoTo errTrap
    Assert.IsNotNothing Nothing
    On Error GoTo 0
    Assert.Fail "IsNotNothing should raise an error."

errTrap:
    Assert.Clear
End Sub
