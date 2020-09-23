VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OneTesFullContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' OneTestFullContext
'
Option Explicit

Public Sub TestFixtureSetup()

End Sub

Public Sub Setup()

End Sub

Public Sub TestOne()
    Assert.Say "hello"
End Sub

Public Sub Teardown()

End Sub

Public Sub TestFixtureTeardown()

End Sub
