VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestATestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestATestCase
'
Option Explicit

Public Sub TestSetOfTestMethods()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(New MockTestCase)
    
    Dim Result As ITestResult
    Set Result = TestCase.Run
    
    Assert.AreEqual 4, Result.TestsRun, "TestSetOfTestMethods should run 4 tests."
End Sub

Public Sub TestFullContextCalled()
    Dim UserTest As New FailableFullContext
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(UserTest)
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.AreEqual 1, Result.TestsRun, "Only 1 test should be run."
    Assert.IsTrue UserTest.TestFixtureSetupCalled, "TestFixtureSetup should be called."
    Assert.IsTrue UserTest.SetupCalled, "Setup should be called."
    Assert.IsTrue UserTest.TearDownCalled, "TearDown should be called."
    Assert.IsTrue UserTest.TestFixtureTearDownCalled, "TestFixtureTearDown should be called."
End Sub

Public Sub TestSetupTearDownCounts()
    Dim UserTest As New CountableSetupTearDown
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(UserTest)
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.AreEqual 2, Result.TestsRun, "2 tests should be run."
    Assert.AreEqual 2, UserTest.SetupCount, "Setup should be called 2 times."
    Assert.AreEqual 2, UserTest.TearDownCount, "TearDown should be called 2 times."
End Sub

Public Sub TestFullContext()
    Dim UserTest As New CountableFullContext
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(UserTest)
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.AreEqual 2, Result.TestsRun, "2 tests should be run."
    Assert.AreEqual 1, UserTest.TestFixtureSetupCount, "TestFixtureSetupCount should be called 1 time."
    Assert.AreEqual 1, UserTest.TestFixtureTearDownCount, "TestFixtureTearDownCount should be called 1 time."
    Assert.AreEqual 2, UserTest.SetupCount, "Setup should be called 2 times."
    Assert.AreEqual 2, UserTest.TearDownCount, "TearDown should be called 2 times."
End Sub

Public Sub TestSort()
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(New MockTestCase)
    
    Test.Sort
    
    Assert.AreEqual "TestFour", Test(1).Name
    Assert.AreEqual "TestOne", Test(2).Name
    Assert.AreEqual "TestThree", Test(3).Name
    Assert.AreEqual "TestTwo", Test(4).Name
End Sub

Public Sub TestFailingTestFixtureSetup()
    Dim UserTest As New FailableFullContext
    UserTest.FailTestFixtureSetup = True
    
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(UserTest)
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.IsTrue Result.IsError, "Test should error."
End Sub


