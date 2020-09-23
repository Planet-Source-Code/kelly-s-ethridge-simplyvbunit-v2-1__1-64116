VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunMyOwnTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub RunOneTest()
    ' We can create a single test to be run.
    ' We create a test method that will call a single
    ' test within our test case class. Any setup/teardown
    ' methods will be included in the test run.
    Dim Test As ITest
    Set Test = Sim.NewTestMethod(New MyOwnTests, "PassingTest")
    
    Dim Result As ITestResult
    Set Result = Test.Run(New DebugWindowListener)
    
    Assert.IsTrue Result.IsSuccess, "My test should pass!"
End Sub

Public Sub RunAnEntireTestCase()
    ' We can create and run an entire test case.
    ' This will run all tests within a test case class.
    Dim Test As ITest
    Set Test = Sim.NewTestCase(New MyOwnTests)
    
    Dim Result As ITestResult
    Set Result = Test.Run(New DebugWindowListener)
    
    ' We know we have three tests in the test case.
    Assert.AreEqual 3, Result.Count, "We should have three test results"
    
    ' Lets get one of the results and inspect it.
    Dim Result1 As ITestResult
    Set Result1 = Result("PassingTest")
    Assert.IsTrue Result1.IsSuccess, "This should be a success!"
    
    ' Now lets deal with the other test result.
    Dim Result2 As ITestResult
    Set Result2 = Result("FailingTest")
    Assert.IsTrue Result.IsFailure, "This should be a failure!"
    
    ' And finally check out the ignored test.
    Dim Result3 As ITestResult
    Set Result3 = Result("IgnoredTest")
    Assert.IsTrue Result3.IsIgnored, "I should have been ignored."
End Sub

Public Sub RunLotsOfTestCases()
    ' We can group our own set of testcases in a test suite
    ' to be run together.
    Dim Suite As New TestSuite
    Suite.Add New FilteredTests
    Suite.Add New MyOwnTests
    
    ' The current Listener and set of Filters are not
    ' applied to the tests to be run. They get their own set.
    Dim Result As ITestResult
    Set Result = Suite.Run(New DebugWindowListener)
    
    Assert.AreEqual 2, Result.Count, "2 test cases should have been run."
    Assert.AreEqual 5, Result.TestsRun, "5 total tests should be run."
End Sub
