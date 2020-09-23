VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSayHear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSayHear
'
Option Explicit

Public Sub TestSay()
    Dim l As New MockTestListener
    Dim Test As TestCase
    Set Test = Sim.NewTestCase(New SayHello)
    
    Assert.Say "Before Run"
    Test.Run l
    Assert.Say "After Run"
    
    Assert.AreEqual 1, l.HearCount, "Hear should be called."
End Sub

