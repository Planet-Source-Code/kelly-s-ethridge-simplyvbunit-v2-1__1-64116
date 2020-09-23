VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMultiCastListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMultiCastListener
'
Option Explicit

Public Sub TestTwoListeners()
    Dim mc As New MultiCastListener
    Dim l1 As New MockTestListener
    Dim l2 As New MockTestListener
    
    mc.Add l1
    mc.Add l2
    
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    
    Suite.Run mc
    
    Assert.AreEqual 1, l1.StartingCount
    Assert.AreEqual 1, l2.StartingCount
    Assert.AreEqual 1, l1.TestSuiteStartingCount
    Assert.AreEqual 1, l2.TestSuiteStartingCount
    Assert.AreEqual 1, l1.TestSuiteCompletedCount
    Assert.AreEqual 1, l2.TestSuiteCompletedCount
    Assert.AreEqual 1, l1.TestCaseStartingCount
    Assert.AreEqual 1, l2.TestCaseStartingCount
    Assert.AreEqual 1, l1.TestCaseCompletedCount
    Assert.AreEqual 1, l2.TestCaseCompletedCount
    Assert.AreEqual 4, l1.TestStartingCount
    Assert.AreEqual 4, l1.TestCompletedCount
    Assert.AreEqual 4, l2.TestStartingCount
    Assert.AreEqual 4, l2.TestCompletedCount
End Sub

Public Sub TestAddRemove()
    Dim mc As New MultiCastListener
    Dim l1 As New MockTestListener
    Dim l2 As New MockTestListener
    
    mc.Add l1
    mc.Add l2
    mc.Remove l1
    
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    
    Suite.Run mc
    
    Assert.AreEqual 0, l1.StartingCount
    Assert.AreEqual 1, l2.StartingCount
    Assert.AreEqual 0, l1.TestSuiteStartingCount
    Assert.AreEqual 1, l2.TestSuiteStartingCount
    Assert.AreEqual 0, l1.TestSuiteCompletedCount
    Assert.AreEqual 1, l2.TestSuiteCompletedCount
    Assert.AreEqual 0, l1.TestCaseStartingCount
    Assert.AreEqual 1, l2.TestCaseStartingCount
    Assert.AreEqual 0, l1.TestCaseCompletedCount
    Assert.AreEqual 1, l2.TestCaseCompletedCount
    Assert.AreEqual 0, l1.TestStartingCount
    Assert.AreEqual 0, l1.TestCompletedCount
    Assert.AreEqual 4, l2.TestStartingCount
    Assert.AreEqual 4, l2.TestCompletedCount
End Sub
