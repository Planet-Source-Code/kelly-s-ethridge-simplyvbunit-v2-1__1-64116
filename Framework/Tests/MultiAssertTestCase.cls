VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiAssertTestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MultiAssertTestCase
'
Option Explicit


' 12 asserts
Public Sub TestOne()
    Assert.IsTrue True
    Assert.IsFalse False
    Assert.AreEqual 1, 1
    Assert.AreNotEqual 1, 2
    Assert.AreSame Me, Me
    Assert.AreNotSame Me, Nothing
    Assert.IsEmpty Empty
    Assert.IsNotEmpty 1
    Assert.IsNothing Nothing
    Assert.IsNotNothing Me
    Assert.IsNull Null
    Assert.IsNotNull 1
End Sub
