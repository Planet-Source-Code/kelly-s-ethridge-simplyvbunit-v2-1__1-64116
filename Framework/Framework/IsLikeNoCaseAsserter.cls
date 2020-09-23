VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsLikeNoCaseAsserter"
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
'    Module: IsLikeNoCaseAsserter
'

Option Explicit
Option Compare Text

Private mAsserter As New TrueAsserter

Friend Function IsLikeNoCase(ByRef Expression As String, ByRef Value As String, ByRef Message As String) As AssertResult
    Set IsLikeNoCase = mAsserter.Assert(Value Like Expression, "IsLikeNoCase", FormatString("Expected {0} to be like {1}", Value, Expression), Message)
End Function

Friend Function IsNotLikeNoCase(ByRef Expression As String, ByRef Value As String, ByRef Message As String) As AssertResult
    Set IsNotLikeNoCase = mAsserter.Assert(Not (Value Like Expression), "IsNotLikeNoCase", FormatString("Expected {0} to not be like {1}", Value, Expression), Message)
End Function
