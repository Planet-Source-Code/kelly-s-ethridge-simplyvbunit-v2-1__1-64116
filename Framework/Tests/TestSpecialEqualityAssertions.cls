VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSpecialEqualityAsserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSpecialEqualityAssertions
'
Option Explicit

Public Sub TestEqual()
    On Error GoTo errTrap
    Assert.AreEqualFloats 1.5, 1.5
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqualFloats should not raise an error."
End Sub

Public Sub TestNotEqual()
    On Error GoTo errTrap
    Assert.AreEqualFloats 1#, 1.00002
    On Error GoTo 0
    Assert.Fail "AreEqualFloats should fail."
    
errTrap:
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error code."
End Sub

Public Sub TestEqualWithinTolerance()
    On Error GoTo errTrap
    Assert.AreEqualFloats 1#, 1.000009
    Assert.AreEqualFloats 1.000009, 1#
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqualFloats should not fail."
End Sub

Public Sub TestEqualStringsSameCase()
    On Error GoTo errTrap
    Assert.AreEqualStrings "hi", "hi"
    Exit Sub

errTrap:
    Assert.Fail "AreEqualStrings should not fail."
End Sub

Public Sub TestEqualStringsDifferentCase()
    On Error GoTo errTrap
    Assert.AreEqualStrings "hi", "HI", vbTextCompare
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqualStrings should not fail."
End Sub

Public Sub TestFailingEqualStringsBinaryCompare()
    On Error GoTo errTrap
    Assert.AreEqualStrings "hi", "Hi", vbBinaryCompare
    On Error GoTo 0
    Assert.Fail "AreEqualStrings should fail."
    
errTrap:
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error code."
End Sub

Public Sub TestFailingEqualStringsTextCompare()
    On Error GoTo errTrap
    Assert.AreEqualStrings "hi", "bye"
    On Error GoTo 0
    Assert.Fail "AreEqualStrings should fail."
    
errTrap:
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error code."
End Sub
