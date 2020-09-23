VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMultiCastFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMultiCastFilter
'
Option Explicit

Public Sub TestCallsMade()
    Dim mf As New MultiCastFilter
    Dim f1 As New MockFilter
    Dim f2 As New MockFilter
    
    mf.Add f1
    mf.Add f2
    
    Dim Test As ITest
    Set Test = New TestSuite
    
    mf.Include Test
    
    Assert.AreEqual 1, f1.IncludeCallCount, "1 call should be made."
    Assert.AreEqual 0, f2.IncludeCallCount, "0 call should be made."
End Sub

Public Sub TestCallsMadeIncudeOnAny()
    Dim mf As New MultiCastFilter
    Dim f1 As New MockFilter
    Dim f2 As New MockFilter
    
    mf.IncludeOnAny = True
    mf.Add f1
    mf.Add f2
    
    Dim Test As ITest
    Set Test = New TestSuite
    
    mf.Include Test
    
    Assert.AreEqual 1, f1.IncludeCallCount, "1 call should be made."
    Assert.AreEqual 1, f2.IncludeCallCount, "1 call should be made."
End Sub

Public Sub TestOneFilter()
    Dim mf As New MultiCastFilter
    Dim f As New NameFilter
    Dim Suite As New TestSuite

    f.Pattern = "*one"
    mf.Add f
    
    Set Suite.Filter = mf
    Suite.Add New MockTestCase

    Dim Result As ITestResult
    Set Result = Suite.Run

    Assert.IsTrue Result.Executed, "Suite should run"
    
    Set Result = Result("MockTestCase")
    Assert.IsTrue Result("TestOne").Executed, "TestOne should run"
    Assert.IsFalse Result("TestTwo").Executed, "TestTwo should not be run"
End Sub

Public Sub TestTwoFiltersAllFail()
    Dim mf As New MultiCastFilter
    Dim f1 As New NameFilter
    Dim f2 As New NameFilter
    Dim Suite As New TestSuite
    
    f1.Pattern = "*One"
    f2.Pattern = "*Two"
    mf.Add f1
    mf.Add f2
    
    Suite.Add New MockTestCase
    Set Suite.Filter = mf
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.IsTrue Result.Executed, "Suite should be run"
    
    Set Result = Result("MockTestCase")
    Assert.IsFalse Result("TestOne").Executed
    Assert.IsFalse Result("TestTwo").Executed
    Assert.IsFalse Result("TestThree").Executed
    Assert.IsFalse Result("TestFour").Executed
End Sub

Public Sub TestTwoFiltersIncludeOnAny()
    Dim mf As New MultiCastFilter
    Dim f1 As New NameFilter
    Dim f2 As New NameFilter
    Dim Suite As New TestSuite
    
    f1.Pattern = "*One"
    f2.Pattern = "*Two"
    mf.Add f1
    mf.Add f2
    mf.IncludeOnAny = True
    
    Suite.Add New MockTestCase
    Set Suite.Filter = mf
    
    Dim Result As ITestResult
    Set Result = Suite.Run
    
    Assert.IsTrue Result.Executed, "Suite should be run"
    
    Set Result = Result("MockTestCase")
    Assert.IsTrue Result("TestOne").Executed
    Assert.IsTrue Result("TestTwo").Executed
    Assert.IsFalse Result("TestThree").Executed
    Assert.IsFalse Result("TestFour").Executed
End Sub
