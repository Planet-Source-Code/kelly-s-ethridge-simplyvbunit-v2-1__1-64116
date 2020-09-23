VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestComparisonAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestComparisonAssertions
'
Option Explicit

Public Sub TestPassingLess()
    On Error GoTo errTrap
    Assert.Less 1, 2
    Assert.Less "A", "a"
    Assert.Less 1#, 1.1
    Assert.Less 1.9, 2
    Assert.Less 2@, 3@
    Exit Sub
    
errTrap:
    
    Assert.Fail "TestPassingLess should not fail."
End Sub

Public Sub TestFailingLessNumbers()
    On Error GoTo errTrap
    Assert.Less 2, 1
    On Error GoTo 0
    Assert.Fail "TestFailingLessNumbers should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number
End Sub

Public Sub TestFailingLessFloats()
    On Error GoTo errTrap
    Assert.Less 1#, 0.9
    On Error GoTo 0
    Assert.Fail "TestFailingLessFloats should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error code."
End Sub

Public Sub TestPassingGreater()
    On Error GoTo errTrap
    Assert.Greater 2, 1
    Assert.Greater "a", "A"
    Assert.Greater 1.1, 1#
    Assert.Greater 2, 1.9
    Assert.Greater 3@, 2@
    Exit Sub

errTrap:
    
    Assert.Fail "TestPassingGreater should not fail."
End Sub

Public Sub TestFailingGreaterNumber()
    On Error GoTo errTrap
    Assert.Greater 1, 2
    On Error GoTo 0
    Assert.Fail "TestFailingGreaterNumber should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error code."
End Sub
