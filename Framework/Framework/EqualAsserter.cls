VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EqualAsserter"
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
'    Module: EqualAsserter
'

''
' An asserter used by several assertions.
'
Option Explicit

Public Expectation As String



Friend Function Assert(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim vt As VbVarType
    vt = VarType(Expected)
    
    ' Since we have so many array types, we will
    ' just bunch them all up into just an array type.
    If CBool(vt And vbArray) Then
        vt = vbArray
    End If
    
    Dim Passed As Boolean
    Select Case vt
        Case vbLong, _
             vbInteger, _
             vbByte, _
             vbCurrency, _
             vbDecimal: Passed = AreEqualNumbers(Expected, Actual)
        Case vbDouble, _
             vbSingle: Passed = AreEqualFloats(Expected, Actual)
        Case vbString:  Passed = AreEqualStrings(Expected, Actual)
        Case vbDate:    Passed = AreEqualDates(Expected, Actual)
        Case vbBoolean: Passed = AreEqualBooleans(Expected, Actual)
        
        Case vbObject, _
             vbDataObject: Passed = AreEqualObjects(Expected, Actual)
        
        Case vbArray
            Passed = AreEqualArrays(Expected, Actual)
    End Select
    
    If (Not Passed) And (vt <> vbArray) Then
        Expectation = GetExpectation(Expected, Actual)
    End If
    
    Assert = Passed
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AreEqualFloats(ByVal Expected As Double, ByRef Actual As Variant) As Boolean
    Select Case VarType(Actual)
        Case vbLong, _
             vbInteger, _
             vbByte, _
             vbDouble, _
             vbSingle, _
             vbCurrency, _
             vbDecimal
             
             AreEqualFloats = (Abs(CDbl(Actual) - Expected) <= modStaticClasses.Assert.FloatCompareTolerance)
    End Select
End Function

Private Function AreEqualObjects(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If IsObject(Actual) Then
        AreEqualObjects = (Expected Is Actual)
    End If
End Function

Private Function AreEqualBooleans(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        AreEqualBooleans = (Expected = Actual)
    End If
End Function

Private Function AreEqualDates(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbDate Then
        AreEqualDates = (DateDiff("s", Expected, Actual) = 0)
    End If
End Function

Private Function AreEqualStrings(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Actual) = vbString Then
        AreEqualStrings = (StrComp(Expected, Actual, modStaticClasses.Assert.StringCompareMethod) = 0)
    End If
End Function

Private Function AreEqualNumbers(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Select Case VarType(Actual)
        Case vbLong, _
             vbInteger, _
             vbByte, _
             vbDouble, _
             vbSingle, _
             vbCurrency, _
             vbDecimal: AreEqualNumbers = (Expected = Actual)
    End Select
End Function

Private Function AreEqualArrays(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    If VarType(Expected) <> VarType(Actual) Then Exit Function
    If LBound(Expected) <> LBound(Actual) Then Exit Function
    If UBound(Expected) <> UBound(Actual) Then Exit Function
    
    Dim Comp As IEqualityComparer
    Set Comp = Comparer.GetComparer((VarType(Expected) And vbObject) = vbObject)
    
    Dim i As Long
    For i = LBound(Expected) To UBound(Expected)
        If Not Comp.Equals(Expected(i), Actual(i)) Then
            Expectation = GetExpectation(Expected(i), Actual(i)) & " at index " & i
            Exit Function
        End If
    Next i

    AreEqualArrays = True
End Function

Private Function GetExpectation(ByRef Expected As Variant, ByRef Actual As Variant) As String
    GetExpectation = FormatString("Expected {0} but was {1}", Expected, Actual)
End Function


