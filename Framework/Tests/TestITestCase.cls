VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestITestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestITestCase
'
Option Explicit

Public Sub TestRunTestOneNoContext()
    Dim Mock As New MockITestCase
    Mock.RunTestOne = True
    
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(Mock)
    
    TestCase.Run
    
    Assert.AreEqual 1, Mock.TestOneRunCount, "TestOne should run"
    Assert.AreEqual 0, Mock.SetupRunCount, "Setup should not run"
    Assert.AreEqual 0, Mock.TeardownRunCount
    Assert.AreEqual 0, Mock.TestFixtureSetupRunCount
    Assert.AreEqual 0, Mock.TestFixtureTeardownRunCount
    Assert.AreEqual 0, Mock.TestTwoRunCount
    Assert.AreEqual 0, Mock.TestThreeRunCount
    Assert.AreEqual 0, Mock.TestFourRunCount
End Sub

Public Sub TestRunSetupTeardown()
    Dim Mock As New MockITestCase
    Mock.RunSetup = True
    Mock.RunTeardown = True
    Mock.RunTestOne = True
    
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(Mock)
    TestCase.Run
    
    Assert.AreEqual 1, Mock.TestOneRunCount, "TestOne should be run"
    Assert.AreEqual 1, Mock.SetupRunCount, "Setup should be run"
    Assert.AreEqual 1, Mock.TeardownRunCount, "Teardown should be run"
End Sub

Public Sub TestRunAll()
    Dim Mock As New MockITestCase
    With Mock
        .RunSetup = True
        .RunTeardown = True
        .RunTestFixtureSetup = True
        .RunTestFixtureTeardown = True
        .RunTestFour = True
        .RunTestOne = True
        .RunTestTwo = True
        .RunTestThree = True
    End With

    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(Mock)

    TestCase.Run

    Assert.AreEqual 1, Mock.TestFixtureSetupRunCount, "TestFixtureSetup should run"
    Assert.AreEqual 4, Mock.SetupRunCount, "TestSetup should run 4 times"
    Assert.AreEqual 1, Mock.TestOneRunCount, "TestOne should run"
    Assert.AreEqual 1, Mock.TestTwoRunCount, "TestTwo should run"
    Assert.AreEqual 1, Mock.TestThreeRunCount, "TestThree should run"
    Assert.AreEqual 1, Mock.TestFourRunCount, "TestFour should run"
    Assert.AreEqual 4, Mock.TeardownRunCount, "TearDown should run 4 times"
    Assert.AreEqual 1, Mock.TestFixtureTeardownRunCount, "TestFixtureTeardown should run"
End Sub

