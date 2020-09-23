VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiCastFilter"
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
'    Module: MultiCaseFilter
'

''
' Allows multiple filters to be applied to tests simultaneously.
'
' @remarks This can be used to combine filters in a complex manner to manage
' which test should be run during test executions.
'
' @see ITestFilter
' @see NameFilter
'
Option Explicit
Implements ITestFilter

Private mFilters        As New Collection
Private mIncludeOnAny   As Boolean



''
' Returns if the test should be included on the first filter that passes it.
'
' @return Returns True if any filter can include a test, False otherwise.
' @remarks By default, all filters must allow the test to be included. Setting
' this to True will allow any single filter to include the test.
'
Public Property Get IncludeOnAny() As Boolean
    IncludeOnAny = mIncludeOnAny
End Property

''
' Sets if the test should included if any filter says to include the test.
'
' @param RHS The new value.
' @remarks By default, all filters must allow the test to be included. Setting
' this to True will allow any single filter to include the test.
'
Public Property Let IncludeOnAny(ByVal RHS As Boolean)
    mIncludeOnAny = RHS
End Property

''
' Adds a filter to a list of filters to be applied to a test.
'
' @param Filter The filter to be added.
'
Public Sub Add(ByVal Filter As ITestFilter)
    If Filter Is Nothing Then
        Call Err.Raise(5, "MultiCastFilter.Add", "Filter cannot be Nothing.")
    End If
    
    On Error GoTo errTrap
    Call mFilters.Add(Filter, CStr(ObjPtr(Filter)))
    Exit Sub
    
errTrap:
    Call Err.Raise(5, , "This filter has already been added.")
End Sub

''
' Removes a filter from the list.
'
' @param Filter The filter to be removed.
'
Public Sub Remove(ByVal Filter As ITestFilter)
    On Error GoTo errTrap
    Call mFilters.Remove(CStr(ObjPtr(Filter)))
errTrap:
End Sub

''
' Returns if the test passed and should be included in the test run.
'
' @param Test The test to apply the filters to.
' @remarks If no filters are in the list then this will return True.
'
Public Function Include(ByVal Test As ITest) As Boolean
    If mFilters.Count = 0 Then
        Include = True
        Exit Function
    End If
    
    Dim Filter As ITestFilter
    For Each Filter In mFilters
        If Filter.Include(Test) Then
            Include = True
            If mIncludeOnAny Then
                Exit Function
            End If
        ElseIf Not mIncludeOnAny Then
            Include = False
            Exit Function
        End If
    Next Filter
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestFilter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestFilter_Include(ByVal Test As ITest) As Boolean
    ITestFilter_Include = Include(Test)
End Function
