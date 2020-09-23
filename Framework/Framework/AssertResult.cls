VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssertResult"
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
'    Module: AssertResult
'

''
' Represents the state of an assertion.
'
' @remarks An AssertResult is used to carry the results and messages
' of an assertion performed. The messages are used to provide the user
' with information about the cause of a failed assertion.
'
Option Explicit

Private mMessage        As String
Private mAssertName     As String
Private mDescription    As String
Private mFailed         As Boolean
Private mErrorCode      As Long
 
 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   We're all friends here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get Message() As String
    Message = mMessage
End Property

Friend Property Get AssertName() As String
    AssertName = mAssertName
End Property

Friend Property Get Description() As String
    Description = mDescription
End Property

Friend Property Get Failed() As Boolean
    Failed = mFailed
End Property

Friend Sub NotifyOnFailure()
    If mFailed Then
        Set Assert.Result = Me
        Call Err.Raise(mErrorCode, mAssertName, mDescription)
    End If
End Sub

Friend Sub SetFailure(ByRef AssertName As String, ByRef Description As String, ByRef Message As String, ByVal ErrorCode As Long)
    mAssertName = AssertName
    mDescription = Description
    mMessage = Message
    mErrorCode = ErrorCode
    mFailed = True
End Sub
