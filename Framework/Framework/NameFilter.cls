VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NameFilter"
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
'    Module: NameFilter
'

''
' Provides a filtering mechanism based on the name and type of test being filtered.
'
' @remarks The framework will use a filter to determine if a test will be included
' or excluded during the run.
' <p>SimplyVBUnit provides one filter type (<b>NameFilter</b>). Other filters
' can be created by implementing the <b>ITestFilter</b> interface.</p>
' <p>The default filtering method is only applied against <b>TestMethod</b> tests.
' <b>TestSuites</b> and <b>TestCases</b> are not filtered out.</p>
'
' @see ITestFilter
' @see MultiCastFilter
'
Option Explicit
Option Compare Text

Implements ITestFilter

Private Const DEF_PATTERN       As String = "*"
Private Const DEF_FILTERMETHODS As Boolean = True

Private mPattern            As String
Private mFilterTestSuites   As Boolean
Private mFilterTestCases    As Boolean
Private mFilterTestMethods  As Boolean
Private mNegate             As Boolean



''
' Returns if the filter is inverting the result.
'
' @return Returns True if the filter is inverting the result, False otherwise.
' @remarks Negating the result is the same as using Not.
'
Public Property Get Negate() As Boolean
    Negate = mNegate
End Property

''
' Sets if the filter should invert the result.
'
' @param RHS The new value.
' @remarks Negating the result is the same as using Not.
'
Public Property Let Negate(ByVal RHS As Boolean)
    mNegate = RHS
End Property

''
' Returns if the filter is evaluating <b>TestSuites</b>
'
' @return Returns True if the filter is evaluating <b>TestSuites</b>.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Get FilterTestSuites() As Boolean
    FilterTestSuites = mFilterTestSuites
End Property

''
' Sets if the filter should evaluate <b>TestSuites</b>
'
' @param RHS Value indicating if <b>TestSuites</b> should be evaluated.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Let FilterTestSuites(ByVal RHS As Boolean)
    mFilterTestSuites = RHS
End Property

''
' Returns if the filter is evaluating <b>TestCases</b>
'
' @return Returns True if the filter is evaluating <b>TestCases</b>.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Get FilterTestCases() As Boolean
    FilterTestCases = mFilterTestCases
End Property
    
''
' Sets if the filter should evaluate <b>TestCases</b>
'
' @param RHS Value indicating if <b>TestCases</b> should be evaluated.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Let FilterTestCases(ByVal RHS As Boolean)
    mFilterTestCases = RHS
End Property

''
' Returns if the filter is evaluating <b>TestMethods</b>
'
' @return Returns True if the filter is evaluating <b>TestMethods</b>.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Get FilterTestMethods() As Boolean
    FilterTestMethods = mFilterTestMethods
End Property

''
' Sets if the filter should evaluate <b>TestMethods</b>
'
' @param RHS Value indicating if <b>TestMethods</b> should be evaluated.
' @Remarks The filter can be set to allow a specific test type
' to be included regardless if it matches the criteria.
'
Public Property Let FilterTestMethods(ByVal RHS As Boolean)
    mFilterTestMethods = RHS
End Property

''
' Returns the pattern used to filter the test names.
'
' @return The pattern used to filter the test names.
' @remarks <b>NameFilter</b> applies the pattern using the Like operator.
' The Option Compare Text is defined so strings are not case sensitive.
'
Public Property Get Pattern() As String
    Pattern = mPattern
End Property

''
' Sets the pattern to be used to filter the test names.
'
' @param RHS The new filter pattern.
' @remarks <b>NameFilter</b> applies the pattern using the Like operator.
' The Option Compare Text is defined so strings are not case sensitive.
'
Public Property Let Pattern(ByVal RHS As String)
    mPattern = RHS
End Property

''
' Performs a filtering operation on the supplied test.
'
' @param Test The test to determine if should be inlcude or excluded from running.
' @return Returns True if the test should be run, False if it should not be run.
'
Public Function Include(ByVal Test As ITest) As Boolean
    If Not mFilterTestSuites Then
        If Test.IsTestSuite Then
            Include = True
            Exit Function
        End If
    End If
    If Not mFilterTestCases Then
        If Test.IsTestCase Then
            Include = True
            Exit Function
        End If
    End If
    If Not mFilterTestMethods Then
        If Test.IsTestMethod Then
            Include = True
            Exit Function
        End If
    End If
    
    Include = (Test.Name Like mPattern)
    If mNegate Then
        Include = Not Include
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Pattern As String, ByVal Negate As Boolean, ByVal FilterTestSuites As Boolean, ByVal FilterTestCases As Boolean, ByVal FilterTestMethods As Boolean)
    mPattern = Pattern
    mNegate = Negate
    mFilterTestSuites = FilterTestSuites
    mFilterTestCases = FilterTestCases
    mFilterTestMethods = FilterTestMethods
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mPattern = DEF_PATTERN
    mFilterTestMethods = DEF_FILTERMETHODS
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestFilter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestFilter_Include(ByVal Test As ITest) As Boolean
    ITestFilter_Include = Include(Test)
End Function
