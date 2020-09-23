VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: SystemEvents
'

Option Explicit

Public Event AddTest(ByVal Test As Object, ByVal Name As String)
Public Event WriteLine(ByVal Text As String)
Public Event AddListener(ByVal Listener As ITestListener)
Public Event AddFilter(ByVal Filter As ITestFilter)
Public Event RemoveFilter(ByVal Filter As ITestFilter)



Friend Sub OnRemoveFilter(ByVal Filter As ITestFilter)
    RaiseEvent RemoveFilter(Filter)
End Sub

Friend Sub OnAddFilter(ByVal Filter As ITestFilter)
    RaiseEvent AddFilter(Filter)
End Sub

Friend Sub OnAddListener(ByVal Listener As ITestListener)
    RaiseEvent AddListener(Listener)
End Sub

Friend Sub OnWriteLine(ByVal Text As String)
    RaiseEvent WriteLine(Text)
End Sub

Friend Sub OnAddTest(ByVal Test As Object, ByVal Name As String)
    RaiseEvent AddTest(Test, Name)
End Sub
