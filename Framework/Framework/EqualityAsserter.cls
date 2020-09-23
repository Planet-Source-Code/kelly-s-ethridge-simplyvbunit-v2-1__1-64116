VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualityAsserter"
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
'    Module: EqualityAsserter
'

Option Explicit

Private mAsserter As New TrueAsserter


Friend Function AreSame(ByRef Expected As IUnknown, ByRef Actual As IUnknown, ByRef message As String) As AssertResult
    Set AreSame = mAsserter.Assert(Expected Is Actual, "AreSame", "<Expected> and <Actual> should reference the same object.", message)
End Function

Friend Function AreNotSame(ByRef NotExpected As IUnknown, ByRef Actual As IUnknown, ByRef message As String) As AssertResult
    Set AreNotSame = mAsserter.Assert(Not (NotExpected Is Actual), "AreNotSame", "<NotExpected> and <Actual> should not be the same object.", message)
End Function

Friend Function AreEqualFloats(ByRef Expected As Double, ByRef Actual As Double, ByVal Tolerance As Double, ByRef message As String) As AssertResult
    Set AreEqualFloats = mAsserter.Assert(Abs(Actual - Expected) <= Tolerance, "AreEqualFloats", "Expected <Actual> to be within " & Tolerance & " of " & Expected & ".", message)
End Function

Friend Function AreEqualStrings(ByRef Expected As String, ByRef Actual As String, ByVal CompareMethod As VbCompareMethod, ByRef message As String)
    Set AreEqualStrings = mAsserter.Assert(StrComp(Expected, Actual, CompareMethod) = 0, "AreEqualStrings", "Expected <" & Expected & "> but was <" & Actual & ">", message)
End Function
