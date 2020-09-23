VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestInEqualityAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestInEqualityAssertions
'
Option Explicit

Public Sub TestNotEqualNumbers()
    On Error GoTo errTrap
    Assert.AreNotEqual 1, 2
    Assert.AreNotEqual 1, "1"
    Exit Sub
    
errTrap:
    
    Assert.Fail "AreNotEqual should not raise an error comparing different values."
End Sub

Public Sub TestFailingNotEqualNumbers()
    On Error GoTo errTrap
    Assert.AreNotEqual 1, 1
    On Error GoTo 0
    Assert.Fail "AreNotEqual should fail with same numbers."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestFailingNotEqualStrings()
    On Error GoTo errTrap
    Assert.AreNotEqual "a", "a"
    On Error GoTo 0
    Assert.Fail "AreNotEqual should fail with same strings."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestPassingDifferentTypes()
    On Error GoTo errTrap
    Assert.AreNotEqual 1, Me
    Exit Sub
    
errTrap:
    
    Assert.Fail "AreNotEqual should not fail with different types."
End Sub
