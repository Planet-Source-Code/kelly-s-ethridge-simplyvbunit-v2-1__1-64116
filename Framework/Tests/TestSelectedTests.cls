VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSelectedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSelectedTests
'
Option Explicit

Public Sub TestUnselectedTestMethod()
    Dim Test As TestMethod
    Dim MyTest As New FailableFullContext
    
    Set Test = Sim.NewTestMethod(MyTest, "TestOne")
    Test.Selected = False
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.IsFalse Result.Executed, "Execute should return false."
    Assert.IsFalse MyTest.TestOneCalled, "TestOne should be called."
End Sub

Public Sub TestUnselectedTestCase()
    Dim Test As TestCase
    Dim MyTest As New FailableFullContext
    
    Set Test = Sim.NewTestCase(MyTest)
    Test.Selected = False
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.IsFalse Result.Executed, "Execute should return false."
    Assert.IsFalse MyTest.SetupCalled, "Setup should not be called."
    Assert.IsFalse MyTest.TestFixtureSetupCalled, "TestFixtureSetup should not be called."
    Assert.IsFalse MyTest.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse MyTest.TearDownCalled, "TearDown should not be called."
    Assert.IsFalse MyTest.TestFixtureTearDownCalled, "TestFixtureTearDown should not be called."
End Sub

Public Sub TestSomeSelectedTests()
    Dim TestCase As TestCase
    Dim MyTest As New MockTestCase

    Set TestCase = Sim.NewTestCase(MyTest)
    
    TestCase("TestOne").Selected = False
    TestCase("TestThree").Selected = False
    TestCase("TestFour").Selected = False
    
    Dim Result As ITestResult
    Set Result = TestCase.Run
    
    Assert.AreEqual 1, Result.TestsRun
    Assert.IsFalse Result("TestOne").Executed, "TestOne should not be executed."
    Assert.IsTrue Result("TestTwo").Executed, "TestTwo should have executed."
    Assert.IsFalse Result("TestThree").Executed, "TestThree should not be executed."
    Assert.IsFalse Result("Testfour").Executed, "TestFour should not be executed."
End Sub

Public Sub TestDeselectAll()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    Suite.DeselectAll
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.IsFalse Result.Executed, "Suite should not run."
    Assert.AreEqual 1, Result.Count, "Result should not have any results."
    
    For Each Result In Result("MockTestCase")
        Assert.IsFalse Result.Executed
    Next Result
End Sub

Public Sub TestDeselectSubTests()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    Suite("MockTestCase").DeselectAll
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.IsTrue Result.Executed
    Assert.AreEqual 1, Result.Count
    Assert.IsFalse Result("MockTestCase").Executed
    
    For Each Result In Result("MockTestCase")
        Assert.IsFalse Result.Executed
    Next Result
End Sub
