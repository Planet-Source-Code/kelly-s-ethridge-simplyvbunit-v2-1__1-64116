Attribute VB_Name = "modPublicFunctions"
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
'    Module: modPublicFunctions
'

Option Explicit

''
' The error code that will be raised by assertions.
'
Public Const ERR_ASSERT_PASS        As Long = 0
Public Const ERR_ASSERT_FAIL        As Long = vbObjectError + 3001
Public Const ERR_ASSERT_IGNORE      As Long = vbObjectError + 3002
Public Const ERR_INVALIDOPERATION   As Long = vbObjectError + 3003
Public Const ERR_METHODNOTFOUND     As Long = 438
Public Const ERR_INVALIDSIGNATURE   As Long = 449

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private mFrequency As Currency



Public Function Listener() As ITestListener
    Set Listener = Iterations.Listener
End Function

Public Function Filters() As FilterList
    Set Filters = Iterations.Filters
End Function

''
' Ensures that a listener is valid.
'
' @param Listener The listener object to ensure that is set to a valid listener.
' @remarks The listener is passed in by reference so it can be set to a valid
' listener if it is currently set to Nothing.
' <p>The default valid listener is a Null listener that simply accepts all
' callbacks, but does not process them in any manner.</p>
'
Public Sub EnsureHaveListener(ByRef Listener As ITestListener)
    If Listener Is Nothing Then
        Set Listener = NullTestListener.NullListener
    End If
End Sub

Public Function FormatString(ByVal s As String, ParamArray Args() As Variant) As String
    Dim i As Long
    For i = 0 To UBound(Args)
        s = Replace$(s, "{" & i & "}", FormatValue(Args(i)))
    Next i
    
    FormatString = s
End Function

Private Function FormatValue(ByRef Value As Variant) As String
    Dim ret As String
    If IsArray(Value) Or IsObject(Value) Then
        ret = TypeName(Value)
    Else
        ret = CStr(Value)
    End If
    
    FormatValue = "<" & ret & ">"
End Function

Public Function GetTicks() As Currency
    If mFrequency = 0@ Then
        If QueryPerformanceFrequency(mFrequency) = 0 Then
            Err.Raise 5, , "Hardware does not support High Performance Counters."
        End If
        
        mFrequency = 1000@ / mFrequency
    End If
    
    Call QueryPerformanceCounter(GetTicks)
End Function

Public Function GetTime(ByRef Ticks As Currency) As Currency
    Dim StopCount As Currency
    Call QueryPerformanceCounter(StopCount)
    GetTime = (StopCount - Ticks) * mFrequency
End Function
