VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CountableFullContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CountableFullContext
'
Option Explicit

Public TestFixtureSetupCount    As Long
Public TestFixtureTearDownCount As Long
Public SetupCount               As Long
Public TearDownCount            As Long



Public Sub TestFixtureSetup()
    TestFixtureSetupCount = TestFixtureSetupCount + 1
End Sub

Public Sub Setup()
    SetupCount = SetupCount + 1
End Sub

Public Sub TearDown()
    TearDownCount = TearDownCount + 1
End Sub

Public Sub TestFixtureTearDown()
    TestFixtureTearDownCount = TestFixtureTearDownCount + 1
End Sub

Public Sub TestOne()

End Sub

Public Sub TestTwo()

End Sub
