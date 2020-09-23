VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsInListAsserter"
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
'    Module: IsInListAsserter
'

Option Explicit

Friend Function IsInList(ByRef Value As Variant, ByRef List() As Variant) As AssertResult
    Dim i As Long
    For i = 0 To UBound(List)
        If EqualAsserter.Assert(Value, List(i)) Then
            Set IsInList = AssertResult.Pass
            Exit Function
        End If
    Next i
    
    Set IsInList = AssertResult.Fail("IsInList", FormatString("Expected to find {0} in list", Value), "Value was not in list.")
End Function

Friend Function IsNotInList(ByRef Value As Variant, ByRef List() As Variant) As AssertResult
    Dim i As Long
    For i = 0 To UBound(List)
        If EqualAsserter.Assert(Value, List(i)) Then
            Set IsNotInList = AssertResult.Fail("IsNotInList", FormatString("Did not expect to find {0} in list", Value), "Value was in the list.")
            Exit Function
        End If
    Next i
    
    Set IsNotInList = AssertResult.Pass
End Function
