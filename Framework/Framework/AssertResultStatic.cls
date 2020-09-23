VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssertResultStatic"
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
'    Module: AssertResultStatic
'

''
' Provides functions for creating new AssertResult objects.
'
Option Explicit

''
' Since all passing assertions have the same state
' we will just use a shared object.
'
' Do not change anything within the default object.
'
Private mPass As New AssertResult


''
' Return an AssertResult that states the assertion passed.
'
' @return The assert result object.
' @remarks The returned object is shared by all methods that
' call this function. The returned result should never be modified.
'
Friend Function Pass() As AssertResult
    Set Pass = mPass
End Function

''
' Returns a new AssertResult indicating an assertion failure.
'
' @param AssertName The name of the assertion that failed.
' @param Expectation The expected assertion to be met.
' @param Message A user readable message describing the failure.
' @return The new result representing an assertion failure.
'
Friend Function Fail(ByRef AssertName As String, ByRef Expectation As String, ByRef Message As String, Optional ByVal ErrorCode As Long = ERR_ASSERT_FAIL) As AssertResult
    Set Fail = New AssertResult
    Call Fail.SetFailure(AssertName, Expectation, Message, ErrorCode)
End Function
