VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompilableTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This unit-test class will work compiled because it implements
' the ITestCase interface.
'
' This interface allows the SimplyVBUnit framework to acquire the
' set of test methods to be called during test execution.
'
Option Explicit
Implements ITestCase

Private Sub ITestCase_AddTests(ByVal Tests As SimplyVBUnitLib.TestCollection)
    Tests.Add "TestOne"
    Tests.Add "TestTwo"
    Tests.Add "Setup"
    
    ' ... Add tests you want
End Sub

Public Sub Setup()
    ' I want a setup called!
    
    Assert.Say "Setup was called!"
End Sub

Public Sub TestOne()
    Assert.Say "This test will run compiled!"
End Sub

Public Sub TestTwo()
    Assert.AreEqual 1, 2, "I hope this failed."
End Sub

Public Sub TestNotIncluded()
    ' This test will not be included ever!
End Sub
