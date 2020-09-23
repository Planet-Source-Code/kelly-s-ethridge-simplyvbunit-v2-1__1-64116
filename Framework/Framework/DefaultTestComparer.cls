VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultTestComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: DefaultTestComparer
'

Option Explicit
Implements ITestComparer

Private Function ITestComparer_Compare(ByVal x As ITest, ByVal y As ITest) As Long
    Dim ret As Long
    
    If x.IsTestSuite Then
        If Not y.IsTestSuite Then
            ret = -1
        End If
    ElseIf x.IsTestCase Then
        If y.IsTestSuite Then
            ret = 1
        ElseIf y.IsTestMethod Then
            ret = -1
        End If
    ElseIf Not y.IsTestMethod Then
        ret = 1
    End If
    
    If ret = 0 Then
        ret = StrComp(x.Name, y.Name, vbTextCompare)
    End If
    
    ITestComparer_Compare = ret
End Function
