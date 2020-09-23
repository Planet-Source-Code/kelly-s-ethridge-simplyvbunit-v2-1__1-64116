VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FailableSetupTearDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SetupTearDownCalled
'
Option Explicit

Public SetupCalled As Boolean
Public TearDownCalled As Boolean
Public FailSetup As Boolean
Public FailTearDown As Boolean
Public TestOneCalled As Boolean

Public Sub Setup()
    SetupCalled = True
    If FailSetup Then
        Err.Raise 5  ' some freaky error
    End If
End Sub

Public Sub TearDown()
    TearDownCalled = True
    If FailTearDown Then
        Err.Raise 5  ' some freaky error
    End If
End Sub

Public Sub TestOne()
    TestOneCalled = True
End Sub


