VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContextResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: ContextResult
'

''
' This is the result of a context method execution.
'
' @remarks Context methods are <b>Setup</b>, <b>Teardown</b>, <b>TestFixtureSetup</b>, and
' <b>TestFixtureTeardown</b>.
'
Option Explicit

Private mMethod         As ContextMethod
Private mIsFailure      As Boolean
Private mIsIgnored      As Boolean
Private mMessage        As String
Private mDescription    As String
Private mCause          As String
Private mExecuted       As Boolean
Private mTest           As Long



''
' Returns a reference to the test that this context method was executed for.
'
' @return The test that was executed that used this context method.
' @remarks This test could be a <b>TestCase</b> or <b>TestMethod</b>.
'
Public Property Get Test() As ITest
    Dim Unk As IUnknown
    Call CopyMemory(Unk, mTest, 4)
    Set Test = Unk
    Call ZeroMemory(Unk, 4)
End Property

''
' Returns the cause of any error that occured.
'
' @return A description of the cause of any error.
'
Public Property Get Cause() As String
    Cause = mCause
End Property

''
' Returns if the context method executed.
'
' @return Returns True if the method executed, False otherwise.
'
Public Property Get Executed() As Boolean
    Executed = mExecuted
End Property

''
' Returns the name of the context method.
'
' @return Returns the name of the context method, not the test name.
'
Public Property Get Name() As String
    Name = mMethod.Name
End Property

''
' Returns if the context method was a failure.
'
' @return Returns True if the context method failed, True otherwise.
'
Public Property Get IsFailure() As Boolean
    IsFailure = mIsFailure
End Property

''
' Returns if the context method was as success.
'
' @return Returns True if the test did not fail and was not ignored.
'
Public Property Get IsSuccess() As Boolean
    If IsFailure Then Exit Property
    If IsIgnored Then Exit Property
    
    IsSuccess = True
End Property

''
' Returns if the context method was ignored.
'
' @return Returns True if the context method was ignored, False otherwise.
'
Public Property Get IsIgnored() As Boolean
    IsIgnored = mIsIgnored
End Property

''
' Returns a message describing the error.
'
' @return The message.
'
Public Property Get Message() As String
    Message = mMessage
End Property

''
' Returns a description about the cause of any problems.
'
' @return A description about the cause of any problems.
'
Public Property Get Description() As String
    Description = mDescription
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Method As ContextMethod, ByVal Test As ITest)
    Set mMethod = Method
    Call SetTest(Test)
End Sub

Friend Sub SetFail(ByVal Result As AssertResult)
    Call InitFromResult(Result)
    mIsFailure = True
End Sub

Friend Sub SetIgnore(ByVal Result As AssertResult)
    Call InitFromResult(Result)
    mIsIgnored = True
End Sub

Friend Sub SetError(ByVal Err As ErrObject, ByVal UserTests As Object)
    Select Case Err.Number
        Case ERR_METHODNOTFOUND
            mMessage = FormatString("Context method {0} not found in class {1}", Name, UserTests)
        Case ERR_INVALIDSIGNATURE
            mMessage = FormatString("Context method {0} in class {1} cannot have parameters", Name, UserTests)
        Case Else
            mMessage = Err.Description
    End Select
    
    mDescription = "Unhandled Error"
    mCause = "Unhandled Error"
    mIsFailure = True
End Sub

Friend Sub SetExecuted()
    mExecuted = True
End Sub

Friend Sub SetTest(ByVal Test As ITest)
    mTest = ObjPtr(Test)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFromResult(ByVal Result As AssertResult)
    mMessage = Result.Message
    mDescription = Result.Description
    mCause = Result.AssertName
End Sub

