VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDifferentAsserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This shows how to use some of the assertions available.
'
Option Explicit

Public Sub TestStringCompareMethod()
    ' we set the string compare method in the Form_Load
    ' so all AreEqual string comparisons will follow that method.
    Assert.AreEqual "hi", "HI", "This should not fail!"
    
    ' We can avoid the comparision method previously set
    ' by using the AreEqualStrings method and providing the method we want to use.
    Assert.AreEqualStrings "hi", "HI", vbBinaryCompare, "This should fail."
End Sub

Public Sub TestFloatCompareTolerance()
    ' we will use the default tolerance of 0.00001
    Assert.AreEqual 9#, 9.000001, "Should be equal!"
    
    ' The Expected value must be a Double or Single
    ' for the tolerance to be used in the comparison.
    Assert.AreEqual 9, 9.000001, "This will fail because we didn't use a float for the expected value."
End Sub

Public Sub TestCatchingAnError()
    On Error GoTo errTrap
    
    ' We have decided that an error is the
    ' correct thing for a test to do and we want to
    ' prevent the error from failing the test.
    
    Err.Raise 5, , "This is correct." ' or this could be a failing assertion.
    
    
errTrap:

    ' Since the error is what we wanted, we must clear
    ' the assertion before exiting or the framework
    ' will become confused and think the test failed.
    Assert.Clear
End Sub

Public Sub TestWithinAList()
    Dim Value As Long
    Value = 1
    
    ' Since we use a ParamArray, there is no place for a message.
    Assert.IsInList Value, 5, 4, 3, 2, 1
End Sub

Public Sub TestAnArrayContainingTheValue()
    Dim Value As Long
    Value = 1
    
    Dim List(4) As Long
    Dim i As Long
    For i = 0 To 4
        List(i) = i
    Next i

    ' This will iterate through the array and use the
    ' same equality comparison as Assert.AreEqual between
    ' each array element and the value searching for.
    Assert.Contains Value, List, "The value should be found in the array."
End Sub

Public Sub TestACollectionContainingTheValue()
    Dim Value As Long
    Value = 1

    Dim Col As New Collection
    Dim i As Long
    For i = 1 To 5
        Col.Add i
    Next i
    
    ' This will iterate through the collection and use the
    ' same equality comparison as Assert.AreEqual between
    ' each collection item and the value searching for.
    Assert.Contains Value, Col, "The collection should contain the value."
End Sub

Public Sub TestYouHaveIgnoredMe()
    ' Ignore will stop the test but will not cause it to fail.
    Assert.Ignore "Why don't you like me?"
    
    Assert.Fail "Should not be called!"
End Sub
