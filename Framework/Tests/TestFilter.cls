VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFilter
'
Option Explicit

Public Sub TestExcludeMethod()
    Dim TestCase As TestCase
    Dim Mock As New MockTestCase
    
    Set TestCase = Sim.NewTestCase(Mock, "TestOne")
    
    Dim Filter As New NameFilter
    Filter.Pattern = "TestOne"
    Filter.Negate = True
    
    Set TestCase.Filter = Filter
    
    Dim Result As ITestResult
    Set Result = TestCase.Run
    
    Assert.IsTrue Result.Executed, "TestCase should have run."
    Assert.IsFalse Result("TestOne").Executed, "TestOne should not have run."
    Assert.IsTrue Result("TestTwo").Executed, "TestTwo should have run."
    Assert.IsTrue Result("TestThree").Executed, "TestThree should have run."
    Assert.IsTrue Result("TestFour").Executed, "TestFour shoudl have run."

    Assert.IsFalse Mock.TestOneCalled
    Assert.IsTrue Mock.TestTwoCalled
    Assert.IsTrue Mock.TestThreeCalled
    Assert.IsTrue Mock.TestFourCalled
End Sub

Public Sub TestExcludeAll()
    Dim TestCase As TestCase
    Dim Mock As New MockTestCase
    
    Set TestCase = Sim.NewTestCase(Mock)
    
    Dim Filter As New NameFilter
    Filter.Pattern = "Test*"
    Filter.Negate = True
    
    Set TestCase.Filter = Filter
    
    Dim Result As ITestResult
    Set Result = TestCase.Run
    
    Assert.IsTrue Result.Executed, "TestCase should have run."
    Assert.IsFalse Result("TestOne").Executed, "TestOne should not have run."
    Assert.IsFalse Result("TestTwo").Executed, "TestTwo should have run."
    Assert.IsFalse Result("TestThree").Executed, "TestThree should have run."
    Assert.IsFalse Result("TestFour").Executed, "TestFour shoudl have run."

    Assert.IsFalse Mock.TestOneCalled
    Assert.IsFalse Mock.TestTwoCalled
    Assert.IsFalse Mock.TestThreeCalled
    Assert.IsFalse Mock.TestFourCalled
End Sub

Public Sub TestIncludeOnlyOne()
    Dim TestCase As TestCase
    Dim Mock As New MockTestCase
    
    Set TestCase = Sim.NewTestCase(Mock)
    
    Dim Filter As New NameFilter
    Filter.Pattern = "*One"
    
    Set TestCase.Filter = Filter
    
    Dim Result As ITestResult
    Set Result = TestCase.Run
    
    Assert.IsTrue Result.Executed, "TestCase should have run."
    Assert.IsTrue Result("TestOne").Executed, "TestOne should have run."
    Assert.IsFalse Result("TestTwo").Executed, "TestTwo should nothave run."
    Assert.IsFalse Result("TestThree").Executed, "TestThree should not have run."
    Assert.IsFalse Result("TestFour").Executed, "TestFour shoudl not have run."

    Assert.IsTrue Mock.TestOneCalled
    Assert.IsFalse Mock.TestTwoCalled
    Assert.IsFalse Mock.TestThreeCalled
    Assert.IsFalse Mock.TestFourCalled
End Sub

Public Sub TestDeepFilter()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase, "1"
    Suite.Add New MockTestCase, "2"
    
    Dim Filter As New NameFilter
    Filter.Pattern = "*Two"
    
    Set Suite.Filter = Filter
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.IsTrue Result.Executed, "Suite should run"
    
    With Result("1")
        Assert.IsTrue .Executed, "TestCase should run"
        Assert.IsTrue .Item("TestTwo").Executed, "TestTwo should run."
        Assert.IsFalse .Item("TestOne").Executed, "TestOne should not run."
        Assert.IsFalse .Item("TestThree").Executed, "TestThree should not run."
        Assert.IsFalse .Item("TestFour").Executed, "TestFour should not run."
    End With
    With Result("2")
        Assert.IsTrue .Executed, "TestCase should run"
        Assert.IsTrue .Item("TestTwo").Executed, "TestTwo should run."
        Assert.IsFalse .Item("TestOne").Executed, "TestOne should not run."
        Assert.IsFalse .Item("TestThree").Executed, "TestThree should not run."
        Assert.IsFalse .Item("TestFour").Executed, "TestFour should not run."
    End With
End Sub
