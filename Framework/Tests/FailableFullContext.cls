VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FailableFullContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FailableFullContext
'
Option Explicit


Public SetupCalled                  As Boolean
Public TearDownCalled               As Boolean
Public TestFixtureSetupCalled       As Boolean
Public TestFixtureTearDownCalled    As Boolean
Public TestOneCalled                As Boolean

Public FailTestFixtureSetup         As Boolean
Public FailTestFixtureTearDown      As Boolean
Public FailTestOne                  As Boolean
Public FailSetup                    As Boolean
Public FailTearDown                 As Boolean


Public Sub Setup()
    SetupCalled = True
    If FailSetup Then
        Err.Raise 5
    End If
End Sub

Public Sub Teardown()
    TearDownCalled = True
    If FailTearDown Then
        Err.Raise 5
    End If
End Sub

Public Sub TestFixtureSetup()
    TestFixtureSetupCalled = True
    If FailTestFixtureSetup Then
        Err.Raise 5
    End If
End Sub
    
Public Sub TestFixtureTeardown()
    TestFixtureTearDownCalled = True
    If FailTestFixtureTearDown Then
        Err.Raise 5
    End If
End Sub

Public Sub TestOne()
    TestOneCalled = True
    If FailTestOne Then
        Err.Raise 5
    End If
End Sub
