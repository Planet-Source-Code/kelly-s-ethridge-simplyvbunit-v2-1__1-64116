VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constructors"
Attribute VB_GlobalNameSpace = True
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
'    Module: Constructors
'

''
' Provides methods to creates specific objects within the framework.
'
' @remarks These methods are globally available and can be called directly.
'
Option Explicit


''
' Returns a new <b>TestMethod</b>.
'
' @param UserTests An object that contains the specified test to be run. This
' will also accept an object that implements the <b>ITestCase</b> interface
' and use that interface to
' @param TestName The name of the test to be run.
' @param RunInContext Selects whether all the setup/teardown code should be included in the test run.
' @return A new <b>TestMethod</b> that can be run directly.
'
Public Function NewTestMethod(ByVal UserTests As Object, ByVal TestName As String, Optional ByVal RunInContext As Boolean = True) As TestMethod
    Set NewTestMethod = New TestMethod
    Call NewTestMethod.Init(UserTests, TestName, RunInContext, Nothing)
End Function

''
' Returns a new <b>TestCase</b>.
'
' @param UserTests An object containing all of the tests to be run.
' @param Name The name of <b>TestCase</b> (The TypeName of the user object is default.)
' @return A new <b>TestCase</b> that can be run directly.
'
Public Function NewTestCase(ByVal UserTests As Object, Optional ByVal Name As String) As TestCase
    Set NewTestCase = New TestCase
    Call NewTestCase.Init(UserTests, Name)
End Function

''
' Returns a new named <b>TestSuite</b>.
'
' @param Name The name of the <b>TestSuite</b>
' @return Returns a named <b>TestSuite</b>
' @remarks A <b>TestSuite</b> can be created directly aswell.
'
Public Function NewTestSuite(ByVal Name As String) As TestSuite
    Set NewTestSuite = New TestSuite
    Call NewTestSuite.Init(Name)
End Function

''
' Returns a new <b>NameFilter</b>.
'
' @param Pattern The pattern the <b>Like</b> operator will use.
' @param Negate Performs a Not operation on the result.
' @param FilterTestSuites Selects whether <b>TestSuite</b>s will be filtered.
' @param FilterTestCases Selects whether <b>TestCase</b>s will be filtered.
' @param FilterTestMethods Selects whether <b>TestMethod</b>s will be filtered.
' @return An initialized <b>NameFilter</b>.
'
Public Function NewNameFilter(ByVal Pattern As String, Optional ByVal Negate As Boolean, Optional ByVal FilterTestSuites As Boolean = False, Optional ByVal FilterTestCases As Boolean = False, Optional ByVal FilterTestMethods As Boolean = True) As NameFilter
    Set NewNameFilter = New NameFilter
    Call NewNameFilter.Init(Pattern, Negate, FilterTestSuites, FilterTestCases, FilterTestMethods)
End Function


''
' for internal use.
'
Friend Function NewTestResult(ByVal Test As ITest) As TestResult
    Set NewTestResult = New TestResult
    Call NewTestResult.Init(Test)
End Function



