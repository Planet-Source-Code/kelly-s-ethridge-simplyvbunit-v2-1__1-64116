VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestIgnore
'
Option Explicit

Public Sub TestIgnoreOneTest()
    Dim UserTests As New IgnorableTestCase
    UserTests.IgnoreTestOne = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(UserTests, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.IsTrue Result.IsIgnored, "Test should be ignored."
    Assert.IsTrue Result.Executed, "Test should be executed."
    Assert.IsFalse Result.IsFailure, "Test should have failed."
End Sub

Public Sub TestIgnoreInTestCase()
    Dim UserTests As New IgnorableTestCase
    UserTests.IgnoreTestOne = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestCase(UserTests)
    
    Dim Result As TestCaseResult
    Set Result = Test.Run
    
    Assert.AreEqual 2, Result.TestsRun, "2 tests should have run."
    Assert.AreEqual 2, Result.Count, "Should have 2 child results."
    Assert.IsFalse Result.IsError, "TestCase should not error."
    
    Dim r As ITestResult
    For Each r In Result
        If r.TestName = "TestOne" Then
            Assert.IsTrue r.IsIgnored, "TestOne should be ignored."
            Assert.IsTrue r.Executed, "TestOne should not be executed."
            Assert.IsFalse r.IsFailure, "TestOne should not have failed."
            Assert.IsFalse r.IsError, "TestOne should not have errored."
        Else
            Assert.IsFalse r.IsIgnored, "TestTwo should not be ignored."
            Assert.IsTrue r.Executed, "TestTwo should be executed."
            Assert.IsFalse r.IsFailure, "TestTwo should not fail."
            Assert.IsFalse r.IsError, "TestTwo should not error."
        End If
    Next r
    
End Sub

Public Sub TestIgnoreInSetup()
    Dim UserTests As New IgnorableTestCase
    UserTests.IgnoreSetup = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestCase(UserTests)
    
    Dim Result As TestCaseResult
    Set Result = Test.Run
    
    Assert.IsFalse UserTests.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse UserTests.TestTwoCalled, "TestTwo should not be called."
    
    Assert.IsFalse Result.IsError, "TestCase should not error."
    Assert.IsFalse Result.IsIgnored, "The TestCase itself should not be ignored."
    Assert.AreEqual 0, Result.TestsRun, "0 tests should be run."
    Assert.AreEqual 2, Result.Count, "2 results should exist."
    
    Dim r As ITestResult
    Set r = Result(1)
    Assert.IsTrue r.IsIgnored, "Test should be ignored."
    Assert.IsFalse r.Executed, "Test should not be executed."
    Assert.IsFalse r.IsFailure, "Test should not fail."
    Assert.IsFalse r.IsError, "Test should not error."
    
    Set r = Result(2)
    Assert.IsTrue r.IsIgnored, "Test should be ignored."
    Assert.IsFalse r.Executed, "Test should not be executed."
    Assert.IsFalse r.IsFailure, "Test should not fail."
    Assert.IsFalse r.IsError, "Test should not error."
End Sub

Public Sub TestIgnoreInTestFixtureSetup()
    Dim UserTests As New IgnorableTestCase
    UserTests.IgnoreTestFixtureSetup = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestCase(UserTests)
    
    Dim Result As TestCaseResult
    Set Result = Test.Run
    
    Assert.IsFalse UserTests.SetupCalled, "Setup should not be called."
    Assert.IsFalse UserTests.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse UserTests.TestTwoCalled, "TestTwo should not be called."
    
    Assert.IsTrue Result.IsIgnored, "TestCase should be ignored."
    Assert.AreEqual 0, Result.TestsRun, "0 tests should be run."
    Assert.AreEqual 0, Result.Count, "0 results should exist."
End Sub
    
