VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContextMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of SimplyVBUnitLib.
'
'    SimplyVBUnitLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: ContextMethod
'

Option Explicit

Private mName       As String
Private mUserTests  As Object



Friend Property Get Name() As String
    Name = mName
End Property

Friend Sub Init(ByVal UserTests As IUnknown, ByVal Name As String)
    Set mUserTests = UserTests
    mName = Name
End Sub

Friend Function Run(ByVal Test As ITest) As ContextResult
    Dim Result As New ContextResult
    Call Result.Init(Me, Test)

    On Error Resume Next
    Call tli.InvokeHookSub(mUserTests, mName, INVOKE_FUNC)
    Select Case Err.Number
        Case ERR_ASSERT_PASS:
        Case ERR_ASSERT_FAIL:   Call Result.SetFail(Assert.Result)
        Case ERR_ASSERT_IGNORE: Call Result.SetIgnore(Assert.Result)
        Case Else:              Call Result.SetError(Err, mUserTests)
    End Select
    On Error GoTo 0

    Call Result.SetExecuted
    Set Run = Result
End Function
