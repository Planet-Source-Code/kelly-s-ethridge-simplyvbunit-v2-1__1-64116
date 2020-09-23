VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConditionAsserter"
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
'    Module: ConditionAsserter
'

''
' Provides several assertions that result in in True or Falst conditions
' using a single value to evaluate.
'
Option Explicit

Private mAsserter As New TrueAsserter


Friend Function IsTrue(ByVal Condition As Boolean, ByRef Message As String) As AssertResult
    Set IsTrue = mAsserter.Assert(Condition, "IsTrue", "Condition should be True", Message)
End Function

Friend Function IsFalse(ByVal Condition As Boolean, ByRef Message As String) As AssertResult
    Set IsFalse = mAsserter.Assert(Not Condition, "IsFalse", "Condition should be False", Message)
End Function

Friend Function IsEmpty(ByRef Value As Variant, ByRef Message As String) As AssertResult
    Set IsEmpty = mAsserter.Assert(VBA.IsEmpty(Value), "IsEmpty", "Value should be Empty", Message)
End Function

Friend Function IsNotEmpty(ByRef Value As Variant, ByRef Message As String) As AssertResult
    Set IsNotEmpty = mAsserter.Assert(Not VBA.IsEmpty(Value), "IsNotEmpty", "Value should not be Empty", Message)
End Function

Friend Function IsNull(ByRef Value As Variant, ByRef Message As String) As AssertResult
    Set IsNull = mAsserter.Assert(VBA.IsNull(Value), "IsNull", "Value should be Null", Message)
End Function

Friend Function IsNotNull(ByRef Value As Variant, ByRef Message As String) As AssertResult
    Set IsNotNull = mAsserter.Assert(Not VBA.IsNull(Value), "IsNotNull", "Value should not be Null", Message)
End Function

Friend Function IsNothing(ByRef Value As IUnknown, ByRef Message As String) As AssertResult
    Set IsNothing = mAsserter.Assert(Value Is Nothing, "IsNothing", "Object variable should be Nothing", Message)
End Function

Friend Function IsNotNothing(ByRef Value As IUnknown, ByRef Message As String) As AssertResult
    Set IsNotNothing = mAsserter.Assert(Not (Value Is Nothing), "IsNotNothing", "Object variable should not be Nothing", Message)
End Function

Friend Function IsLike(ByRef Expression As String, ByRef Value As String, ByRef Message As String) As AssertResult
    Set IsLike = mAsserter.Assert(Value Like Expression, "IsLike", "Expected <" & Value & "> to be like <" & Expression & ">", Message)
End Function

Friend Function IsNotLike(ByRef Expression As String, ByRef Value As String, ByRef Message As String) As AssertResult
    Set IsNotLike = mAsserter.Assert(Not (Value Like Expression), "IsNotLike", "Expected <" & Value & "> to not be like <" & Expression & ">", Message)
End Function
