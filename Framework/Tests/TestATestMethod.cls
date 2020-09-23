VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestATestMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestATestMethod
'
Option Explicit

Public Sub TestMethodCalled()
    Dim Tests As New TestSuiteRunsTests
    Dim Test As TestMethod
    Set Test = NewTestMethod(Tests, "TestOne")
    Test.Run
    Assert.IsTrue Tests.TestOneWasRun, "TestOne should have been run."
End Sub

Public Sub TestSuccessFulResult()
    Dim Tests As New TestSuiteRunsTests
    Dim Test As TestMethod
    Set Test = NewTestMethod(Tests, "TestOne")
    
    Dim Result As TestResult
    Set Result = Test.Run
    Assert.IsNotNothing Result, "Result shoult not be Nothing."
    Assert.IsTrue Result.IsSuccess, "Result should be successful."
    Assert.IsFalse Result.IsFailure, "Result should not be a failure."
    Assert.AreEqual "TestOne", Result.TestName, "Name should be 'TestOne'."
End Sub

Public Sub TestFailingResult()
    Dim Tests As New FailingTestCase
    Dim Test As TestMethod
    Set Test = NewTestMethod(Tests, "IWillFail")
    
    Dim Result As TestResult
    Set Result = Test.Run
    Assert.IsTrue Result.IsFailure, "Result should be a failure."
    Assert.IsFalse Result.IsSuccess, "Result should not be a success."
    Assert.AreEqual "IWillFail", Result.TestName
    Assert.AreEqual "I Will Fail", Result.Message
End Sub

Public Sub TestAssertCount()
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(New MultiAssertTestCase, "TestOne")
    Dim Result As ITestResult
    Set Result = Test.Run

    Assert.AreEqual 12, Result.AssertCount, "AssertCount should return 12."

End Sub

Public Sub TestSetupTearDown()
    Dim Test As ITest
    Dim UserTest As New FailableSetupTearDown
    Dim Result As ITestResult
    
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Set Result = Test.Run
    
    Assert.IsTrue UserTest.SetupCalled, "Setup should be called."
    Assert.IsTrue UserTest.TearDownCalled, "TearDown should be called."
    Assert.AreEqual 1, Result.TestsRun, "Only 1 test should be run."
End Sub

Public Sub TestFailingSetup()
    Dim UserTest As New FailableSetupTearDown
    UserTest.FailSetup = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.AreEqual 0, Result.TestsRun, "No tests should be run."
    Assert.IsTrue UserTest.SetupCalled, "SetupShould be called."
    Assert.IsFalse UserTest.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse UserTest.TearDownCalled, "TearDown should not be called."
End Sub

Public Sub TestFailingTearDown()
    Dim UserTest As New FailableSetupTearDown
    UserTest.FailTearDown = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.AreEqual 1, Result.TestsRun, "1 test should be run."
    Assert.IsTrue UserTest.SetupCalled, "Setup should be called."
    Assert.IsTrue UserTest.TearDownCalled, "TearDown should be called."
    Assert.IsTrue Result.IsError, "Test should error."
End Sub

Public Sub TestFailingTestFixtureSetupInFullContext()
    Dim UserTest As New FailableFullContext
    UserTest.FailTestFixtureSetup = True
    
    Dim Test As TestMethod
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.IsTrue Result.IsError, "Test should have errored."
    Assert.IsTrue UserTest.TestFixtureSetupCalled, "TestFixtureSetup should be called."
    Assert.IsFalse UserTest.SetupCalled, "Setup should not be called."
    Assert.IsFalse UserTest.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse UserTest.TearDownCalled, "TearDown should not be called."
    Assert.IsFalse UserTest.TestFixtureTearDownCalled, "TestFixtureTearDown should not be called."
End Sub

Public Sub TestFailingSetupInFullContext()
    Dim UserTest As New FailableFullContext
    UserTest.FailSetup = True
    
    Dim Test As TestMethod
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.IsTrue Result.IsError, "Test should have error."
    Assert.IsTrue UserTest.TestFixtureSetupCalled, "TestFixtureSetup should be called."
    Assert.IsTrue UserTest.SetupCalled, "Setup should be called."
    Assert.IsFalse UserTest.TestOneCalled, "TestOne should not be called."
    Assert.IsFalse UserTest.TearDownCalled, "TearDown should not be called."
    Assert.IsTrue UserTest.TestFixtureTearDownCalled, "TestFixtureTearDown should be called."
End Sub

Public Sub TestFailingTearDownInFullContext()
    Dim UserTest As New FailableFullContext
    UserTest.FailTearDown = True
    
    Dim Test As TestMethod
    Set Test = Sim.NewTestMethod(UserTest, "TestOne")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    Assert.IsTrue Result.IsError, "Test should have errored."
    Assert.AreEqual "TestOne", Result.TestName, "TestName should be 'TestOne'."
    Assert.IsTrue UserTest.TestFixtureSetupCalled, "TestFixtureSetup should be called."
    Assert.IsTrue UserTest.SetupCalled, "Setup should be called."
    Assert.IsTrue UserTest.TestOneCalled, "TestOne should be called."
    Assert.IsTrue UserTest.TearDownCalled, "TearDown should be called."
    Assert.IsTrue UserTest.TestFixtureTearDownCalled, "TestFixtureTearDown should be called."
End Sub

Public Sub TestCallFullContextFromTestCaseChild()
    Dim TestCase As TestCase
    Dim User As New CountableFullContext
    Set TestCase = Sim.NewTestCase(User)
    
    Dim Method As TestMethod
    Set Method = TestCase("TestOne")
    
    Method.Run
    
    Assert.AreEqual 1, User.TestFixtureSetupCount, "TestFixtureSetup should have run."
    Assert.AreEqual 1, User.SetupCount, "Setup should have run."
    Assert.AreEqual 1, User.TearDownCount, "Teardown should have run."
    Assert.AreEqual 1, User.TestFixtureTearDownCount, "TestFixtureTeardown should have run."
End Sub

Public Sub TestGetResultWithTestFixtureIgnore()
    Dim TestCase As TestCase
    Dim User As New IgnorableTestCase
    Set TestCase = Sim.NewTestCase(User)
    User.IgnoreTestFixtureSetup = True
    
    Dim Method As TestMethod
    Set Method = TestCase("TestOne")
    
    On Error Resume Next
    Dim Result As ITestResult
    Set Result = Method.Run
    On Error GoTo 0
    
    Assert.IsNotNothing Result
End Sub

Public Sub TestMethodNotFound()
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(New MockTestCase, "MissingTest")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.IsTrue Result.IsError, "Result should report an error."
End Sub

Public Sub TestRunInvalidTest()
    Dim Mock As New MockITestCase
    Mock.RunInvalidTest = True
    
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(Mock, "TestWithParameter")
    
    Dim Result As ITestResult
    Set Result = Test.Run
    
    Assert.IsTrue Result.IsError, "Result should report an error."
End Sub
