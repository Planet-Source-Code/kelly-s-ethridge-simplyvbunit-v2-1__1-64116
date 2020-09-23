VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestATestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCascadingTestSuites
'
Option Explicit

Public Sub TestTwoSuites()
    Dim outer As New TestSuite
    Dim inner As New TestSuite
    
    inner.Add New MockTestCase
    outer.Add New MockTestCase
    outer.Add inner
    
    Dim Listener As New MockTestListener
    Dim Results As TestSuiteResult
    Set Results = outer.Run(Listener)
    
    Assert.AreEqual 8, Results.TestsRun, "Results should contain 8 test results."
    
    Dim SuiteResults As Long
    Dim TestResults As Long
    
    Dim Result As ITestResult
    For Each Result In Results
        If TypeOf Result Is TestSuiteResult Then
            SuiteResults = SuiteResults + 1
        Else
            TestResults = TestResults + 1
        End If
    Next Result
    
    Assert.AreEqual 1, SuiteResults, "There should be only 1 TestSuiteResult."
    Assert.AreEqual 1, TestResults, "There should be 4 TestResults."
    Assert.AreEqual 1, Listener.StartingCount, "Listener.Starting should be called 1 time."
    Assert.AreEqual "\TestSuite", outer.FullName
    Assert.AreEqual "\TestSuite\TestSuite", inner.FullName
End Sub

Public Sub TestAssertCount()
    Dim Suite As New TestSuite
    Suite.Add New MultiAssertTestCase, "MultiAssertTestCase1"
    Suite.Add New MultiAssertTestCase, "MultiAssertTestCase2"
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.AreEqual 24, Result.AssertCount, "TestAssertCount should have 24 assertions."
End Sub

Public Sub TestDeepAssertCount()
    Dim Suite1 As New TestSuite
    Suite1.Add New MultiAssertTestCase, "1"
    Suite1.Add New MultiAssertTestCase, "2"
    Suite1.Add Sim.NewTestMethod(New MockTestCase, "TestThree")
    
    Dim Suite2 As New TestSuite
    Suite2.Add New MultiAssertTestCase
    Suite2.Add Suite1
    Suite2.Add Sim.NewTestMethod(New MockTestCase, "TestFour")
    
    Dim Result As ITestResult
    Set Result = Suite2.Run
    
    Assert.AreEqual 38, Result.AssertCount, "TestAssertcount should have 38 assertions."
End Sub

Public Sub TestListenerCallbackSequence()
    Dim Suite As New TestSuite
    Suite.Add New OneTesFullContext
    
    Dim l As New MockListenerWithOrder
    Suite.Run l
    
    Assert.AreEqual 1, l.RunStarted
    Assert.AreEqual 2, l.TestSuiteStarted
    Assert.AreEqual 3, l.TestCaseStarted
    Assert.AreEqual 4, l.TestFixtureSetupStarted
    Assert.AreEqual 5, l.TestFixtureSetupCompleted
    Assert.AreEqual 6, l.SetupStarted
    Assert.AreEqual 7, l.SetupCompleted
    Assert.AreEqual 8, l.TestStarted
    Assert.AreEqual 9, l.Hear
    Assert.AreEqual 10, l.TestCompleted
    Assert.AreEqual 11, l.TeardownStarted
    Assert.AreEqual 12, l.TeardownCompleted
    Assert.AreEqual 13, l.TestFixtureTeardownStarted
    Assert.AreEqual 14, l.TestFixtureTeardownCompleted
    Assert.AreEqual 15, l.TestCaseCompleted
    Assert.AreEqual 16, l.TestSuiteCompleted
    Assert.AreEqual 17, l.RunCompleted
End Sub

Public Sub TestTotalTestCount()
    Dim s1 As New TestSuite
    s1.Add New FailingTestCase
    s1.Add New CountableSetupTearDown
    
    Dim s2 As New TestSuite
    s2.Add New CountableSetupTearDown
    s1.Add s2
    
    Assert.AreEqual 5, s1.TotalTestCount
End Sub
