VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreEqualDatesAsserter"
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
'    Module: AreEqualDatesAsserter
'

Option Explicit

Private mPrecisions()       As String
Private mPrecisionNames()   As String
Private mAsserter           As New TrueAsserter



Friend Function Assert(ByVal Expected As Date, ByVal Actual As Date, ByVal Precision As DateComparePrecision, ByRef Message As String) As AssertResult
    Dim Diff As Long
    Diff = Abs(DateDiff(mPrecisions(Precision), Expected, Actual))
    Set Assert = mAsserter.Assert(CBool(Diff = 0), "AreEqualDates", GetExpectation(Expected, Actual, Precision), Message)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitPrecisions()
    ReDim mPrecisions(0 To 8)
    Call SetPrecision(ToTheDay, "d")
    Call SetPrecision(ToTheDayOfYear, "y")
    Call SetPrecision(ToTheHour, "h")
    Call SetPrecision(ToTheMinute, "n")
    Call SetPrecision(ToTheMonth, "m")
    Call SetPrecision(ToTheQuarter, "q")
    Call SetPrecision(ToTheSecond, "s")
    Call SetPrecision(ToTheWeek, "ww")
    Call SetPrecision(ToTheYear, "yyyy")
End Sub

Private Sub SetPrecision(ByVal Precision As DateComparePrecision, ByVal PrecisionString As String)
    mPrecisions(Precision) = PrecisionString
End Sub

Private Sub InitPrecisionNames()
    ReDim mPrecisionNames(0 To 8)
    Call SetName(ToTheDay, "Day")
    Call SetName(ToTheDayOfYear, "Day of the year")
    Call SetName(ToTheHour, "Hour")
    Call SetName(ToTheMinute, "Minute")
    Call SetName(ToTheMonth, "Month")
    Call SetName(ToTheQuarter, "Quarter")
    Call SetName(ToTheSecond, "Second")
    Call SetName(ToTheWeek, "Week")
    Call SetName(ToTheYear, "Year")
End Sub

Private Sub SetName(ByVal Precision As DateComparePrecision, ByVal Name As String)
    mPrecisionNames(Precision) = Name
End Sub

Private Function GetExpectation(ByVal Expected As Date, ByVal Actual As Date, ByVal Precision As DateComparePrecision) As String
    GetExpectation = FormatString("{0} should be equal to the same {1} as {2}", Actual, mPrecisionNames(Precision), Expected)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitPrecisions
    Call InitPrecisionNames
End Sub
