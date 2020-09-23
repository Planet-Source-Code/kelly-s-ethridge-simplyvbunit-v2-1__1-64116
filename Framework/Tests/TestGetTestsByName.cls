VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestGetTestsByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestGetTestsByName
'
Option Explicit

Public Sub TestInSuite()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase, "TestCase1"
    Suite.Add New MockTestCase, "TestCase2"
    Suite.Add New MockTestCase, "TestCase3"
    
    Dim Test As ITest
    Set Test = Suite("TestCase1")
    Assert.IsNotNothing Test, "Test should be Test1"
    Assert.AreEqual Test.Name, "TestCase1", "Name should be TestCase1"
    
    Set Test = Suite("TestCase2")
    Assert.IsNotNothing Test, "Test should be Test2"
    Assert.AreEqual Test.Name, "TestCase2", "Name should be TestCase2"
    
    Set Test = Suite("TestCase3")
    Assert.IsNotNothing Test, "Test should be Test3"
    Assert.AreEqual Test.Name, "TestCase3", "Name should be TestCase3"
End Sub

Public Sub TestInSuiteUnnamedTestCases()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    Suite.Add New FailableFullContext
    
    Dim Test As ITest
    Set Test = Suite("MockTestCase")
    Assert.IsNotNothing Test, "Test should be MockTestCase"
    Assert.AreEqual "MockTestCase", Test.Name, "Test should be named MockTestCase"
    
    Set Test = Suite("FailableFullContext")
    Assert.IsNotNothing Test, "Test should be FailableFullContext"
    Assert.AreEqual "FailableFullContext", Test.Name, "test name should be 'FailableFullContext'"
End Sub

Public Sub TestInTestCase()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(New MockTestCase)
    
    Dim Test As ITest
    Set Test = TestCase("testone")
    Assert.IsNotNothing Test, "Test should be TestOne"
    Assert.AreEqual "TestOne", Test.Name, "Test name should be 'TestOne'"
End Sub

Public Sub TestFindInSuite()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    
    Dim Test As ITest
    Set Test = Suite.Find("MockTestCase")
    Assert.IsNotNothing Test, "Test should be MockTestCase"
    Assert.AreEqual "MockTestCase", Test.Name, "Test name should be MockTestCase"
    
    Set Test = Suite.Find("TestTwo")
    Assert.IsNotNothing Test, "Test should be TestTwo"
    Assert.AreEqual "TestTwo", Test.Name, "Test name should be 'TestTwo"
End Sub

Public Sub TestFindInTestCase()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(New MockTestCase)
    
    Dim Test As ITest
    Set Test = TestCase.Find("TestThree")
    Assert.IsNotNothing Test
    Assert.AreEqual "TestThree", Test.Name
End Sub

Public Sub TestNotFoundInTestCase()
    Dim TestCase As TestCase
    Set TestCase = Sim.NewTestCase(New MockTestCase)
    
    Dim Test As ITest
    On Error GoTo errTrap
    Set Test = TestCase.Find("ABC")
    Assert.IsNothing Test
    Exit Sub
    
errTrap:
    Assert.Fail "TestNotFoundInTestCase should not fail."
End Sub

Public Sub TestNotFoundInSuite()
    Dim Suite As New TestSuite
    Suite.Add New MockTestCase
    Suite.Add New FailableFullContext
    
    Dim Test As ITest
    On Error GoTo errTrap
    Set Test = Suite.Find("Abc")
    Assert.IsNothing Test
    Exit Sub
    
errTrap:
    Assert.Fail "TestNotFoundInSuite should not fail."
End Sub
