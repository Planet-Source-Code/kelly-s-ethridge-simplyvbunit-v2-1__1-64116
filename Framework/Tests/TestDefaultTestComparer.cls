VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDefaultTestComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDefaultTestComparer
'
Option Explicit

Private Comp As ITestComparer
Private s1 As TestSuite


Public Sub Setup()
    Set Comp = New DefaultTestComparer
End Sub

Public Sub TestSuites()
    Dim s1 As TestSuite
    Dim s2 As TestSuite
    
    Set s1 = Sim.NewTestSuite("a")
    Set s2 = Sim.NewTestSuite("b")
    
    Assert.Less Comp.Compare(s1, s2), 0, "s1 should be less than s2"
    Assert.Greater Comp.Compare(s2, s1), 0, "s2 should be greater than s1"
    Assert.AreEqual 0, Comp.Compare(s1, s1), "s1 should equal itself"
End Sub

Public Sub TestTestCases()
    Dim t1 As TestCase
    Dim t2 As TestCase
    
    Set t1 = Sim.NewTestCase(New MockITestCase, "a")
    Set t2 = Sim.NewTestCase(New MockITestCase, "b")
    
    Assert.Less Comp.Compare(t1, t2), 0, "t1 should be less than t2"
    Assert.Greater Comp.Compare(t2, t1), 0, "t2 should be greater than t1"
    Assert.AreEqual 0, Comp.Compare(t1, t1), "t1 should equal itself"
End Sub

Public Sub TestTestMethods()
    Dim m1 As TestMethod
    Dim m2 As TestMethod
    
    Set m1 = Sim.NewTestMethod(New MockITestCase, "TestOne")
    Set m2 = Sim.NewTestMethod(New MockITestCase, "TestTwo")
    
    Assert.Less Comp.Compare(m1, m2), 0, "m1 should be less than m2"
    Assert.Greater Comp.Compare(m2, m1), 0, "m2 should be greater than m1"
    Assert.AreEqual 0, Comp.Compare(m1, m1), "m1 should equal itself"
End Sub

Public Sub TestSuiteAndCase()
    Dim s1 As TestSuite
    Dim t1 As TestCase
    
    Set s1 = Sim.NewTestSuite("z")
    Set t1 = Sim.NewTestCase(New MockITestCase, "a")
    
    Assert.Less Comp.Compare(s1, t1), 0, "TestSuite should be less than TestCase"
    Assert.Greater Comp.Compare(t1, s1), 0, "TestCase should be greater than TestSuite"
End Sub

Public Sub TestSuiteAndMethod()
    Dim s1 As TestSuite
    Dim m1 As TestMethod
    
    Set s1 = Sim.NewTestSuite("z")
    Set m1 = Sim.NewTestMethod(New MockITestCase, "TestOne")
    
    Assert.Less Comp.Compare(s1, m1), 0, "TestSuite should be less than TestMethod"
    Assert.Greater Comp.Compare(m1, s1), 0, "TestMethod should be greater than TestSuite"
End Sub

Public Sub TestCaseAndMethod()
    Dim t1 As TestCase
    Dim m1 As TestMethod
    
    Set t1 = Sim.NewTestCase(New MockITestCase, "z")
    Set m1 = Sim.NewTestMethod(New MockITestCase, "TestOne")
    
    Assert.Less Comp.Compare(t1, m1), 0, "TestCase should be less than TestMethod"
    Assert.Greater Comp.Compare(m1, t1), 0, "TestMethod should be greater than TestCase"
End Sub
