VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IgnorableTestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' IgnorableTestCase
'
Option Explicit

Public TestOneCalled            As Boolean
Public TestTwoCalled            As Boolean
Public SetupCalled              As Boolean
Public TestFixtureSetupCalled   As Boolean

Public IgnoreTestOne            As Boolean
Public IgnoreTestTwo            As Boolean
Public IgnoreSetup              As Boolean
Public IgnoreTestFixtureSetup   As Boolean


Public Sub TestFixtureSetup()
    TestFixtureSetupCalled = True
    If IgnoreTestFixtureSetup Then
        Assert.Ignore
    End If
End Sub

Public Sub Setup()
    SetupCalled = True
    If IgnoreSetup Then
        Assert.Ignore
    End If
End Sub

Public Sub TestOne()
    TestOneCalled = True
    If IgnoreTestOne Then
        Assert.Ignore
    End If
End Sub

Public Sub TestTwo()
    TestTwoCalled = True
    If IgnoreTestTwo Then
        Assert.Ignore
    End If
End Sub

