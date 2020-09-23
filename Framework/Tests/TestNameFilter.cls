VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNameFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestNameFilter
'
Option Explicit

Private f As NameFilter
Private Abc As ITest
Private Def As ITest
Private aDE As ITest

Public Sub Setup()
    Set f = New NameFilter
    f.FilterTestSuites = True
    
    Set Abc = Sim.NewTestSuite("Abc")
    Set Def = Sim.NewTestSuite("def")
    Set aDE = Sim.NewTestSuite("aDE")
End Sub

Public Sub Default()
    Assert.IsTrue f.Include(Abc)
    Assert.IsTrue f.Include(Def)
    Assert.IsTrue f.Include(aDE)
End Sub

Public Sub TestFilterAbc()
    f.Pattern = "ABC*"
    
    Assert.IsTrue f.Include(Abc)
    Assert.IsFalse f.Include(Def)
    Assert.IsFalse f.Include(aDE)
End Sub

Public Sub TestFilterAnyDEAny()
    f.Pattern = "*de*"
    
    Assert.IsFalse f.Include(Abc)
    Assert.IsTrue f.Include(Def)
    Assert.IsTrue f.Include(aDE)
End Sub

Public Sub TestFilterAnyDE()
    f.Pattern = "*de"
    
    Assert.IsFalse f.Include(Abc)
    Assert.IsFalse f.Include(Def)
    Assert.IsTrue f.Include(aDE)
End Sub

Public Sub TestNegate()
    f.Pattern = "Abc"
    f.Negate = True
    
    Assert.IsFalse f.Include(Abc)
    Assert.IsTrue f.Include(Def)
    Assert.IsTrue f.Include(aDE)
End Sub


