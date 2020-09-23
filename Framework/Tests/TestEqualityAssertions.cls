VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEqualityAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEqualityAssertions
'
Option Explicit

Public Sub TestEqualNumbers()
    On Error GoTo errTrap
    Assert.AreEqual 1, 1&
    Assert.AreEqual 1&, 1
    Assert.AreEqual 1#, 1!
    Assert.AreEqual 1!, 1#
    Assert.AreEqual CDec(1), 1@
    Assert.AreEqual 1@, CDec(1)
    Assert.AreEqual 1#, 1
    Assert.AreEqual 1, 1#
    Assert.AreEqual 1@, CByte(1)
    Assert.AreEqual CByte(1), 1@
    Exit Sub
    
errTrap:
    Assert.Fail Err.Description & ": AreEqual should not raise an error."
End Sub

Public Sub TestUnEqualNumbers()
    On Error GoTo errTrap
    Assert.AreEqual 1, 2
    On Error GoTo 0
    Assert.Fail "AreEqual should raise an error."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestFailingAreEqualNumberAndString()
    On Error GoTo Continue:
    Assert.AreEqual 1, "1"
    On Error GoTo 0
    Assert.Fail "AreEqual should not compare numbers to strings."
    
Continue:
    On Error GoTo -1
    On Error GoTo errTrap
    Assert.AreEqual "1", 1
    On Error GoTo 0
    Assert.Fail "AreEqual should not compare numbers to strings."

errTrap:
    Assert.Clear
End Sub

Public Sub TestFailingAreEqualNumberAndArray()
    Dim a(0) As Long
    On Error GoTo errTrap
    Assert.AreEqual 1, a
    On Error GoTo 0
    Assert.Fail "AreEqual should not be able to compare a number and array."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "AreEqual raised the wrong error code."
End Sub

Public Sub TestEqualArrays()
    On Error GoTo errTrap
    Dim a1() As String
    Dim a2() As String
    
    ReDim a1(10)
    Dim i As Long
    For i = 0 To 10
        a1(i) = CStr(i)
    Next i
    a2 = a1
    
    Assert.AreEqual a1, a2
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqual should not raise an error."
End Sub

Public Sub TestEqualStrings()
    On Error GoTo errTrap
    Assert.AreEqual "hello", "hello"
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqual should compare eqaul strings."
End Sub

Public Sub TestNotEqualStrings()
    On Error GoTo errTrap
    Assert.AreEqual "hello", "bye"
    On Error GoTo 0
    Assert.Fail "AreEqual should fail different strings."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestEqualDates()
    On Error GoTo errTrap
    Assert.AreEqual #1/1/2001#, #1/1/2001#
    Exit Sub
    
errTrap:
    Assert.Fail "AreEqual should compare equal dates."
End Sub

Public Sub TestNotEqualDates()
    On Error GoTo errTrap
    Assert.AreEqual #1/1/2001#, #1/2/2001#
    On Error GoTo 0
    Assert.Fail "AreEqual should fail different dates."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestCompareDateToStringDate()
    On Error GoTo errTrap
    Assert.AreEqual #1/1/2001#, "1/1/2001"
    On Error GoTo 0
    Assert.Fail "AreEqual should fail dates and strings."
    
errTrap:
    Assert.Clear
End Sub

Public Sub TestWithinDefaultTolerance()
    On Error GoTo errTrap
    Assert.AreEqual 1#, 1.000009
    Exit Sub
    
errTrap:
    Assert.Fail "Should be within tolerance."
End Sub

