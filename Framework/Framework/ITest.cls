VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: ITest
'

''
' Provides a common interface used by the three(3) test types in the framework.
'
' @remarks Most of the functionality among the test types are common and can
' be utilizided the same way. This interface provides a common method for accessing
' the test information regardless of the type of test.
'
' @see TestSuite
' @see TestCase
' @see TestMethod
'
Option Explicit


Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Runs the specific test or set of tests.
'
' @param Listener An object used to receive callbacks of specific events during a test run.
' @return Returns the result of the test or tests being run.
'
Public Function Run(Optional ByVal Listener As ITestListener) As ITestResult: End Function

''
' Returns the name of the test instance.
'
' @return Returns the name of this test.
'
Public Property Get Name() As String: End Property

''
' Returns if this specific test instance is a <b>TestSuite</b>.
'
' @return Returns True if this object is a <b>TestSuite</b>, False otherwise.
'
Public Property Get IsTestSuite() As Boolean: End Property

''
' Returns if the specific test instance is a <b>TestCase</b>.
'
' @return Returns True if this object is a <b>TestCase</b>, False otherwise.
'
Public Property Get IsTestCase() As Boolean: End Property

''
' Returns if the specific test instance is a <b>TestMethod</b>.
'
' @return Returns True if the object is a <b>TestMethod</b>, False otherwise.
'
Public Property Get IsTestMethod() As Boolean: End Property

''
' Returns the number of tests contained within this test at the first level.
'
' @return Returns the number of test contained in this test at the first level.
'
Public Property Get Count() As Long: End Property

''
' Returns the parent of this test if it has one.
'
' @return The parent of this test is returned, or Nothing if no parent exists.
'
Public Property Get Parent() As ITest: End Property

''
' Sets the parent for this test.
'
' @param RHS The parent to be set for this test.
'
Public Property Set Parent(ByVal RHS As ITest): End Property

''
' Returns if this test has a parent.
'
' @return Returns True if a parent exists, False otherwise.
'
Public Property Get HasParent() As Boolean: End Property

''
' Searches for a test based on the name.
'
' @param Name The name of the test being searched for.
' @return Returns the test if found, Nothing if not found.
' @remarks Since multiple tests may have the same name, this will
' generally return the first match found.
'
Public Function Find(ByVal Name As String) As ITest: End Function

''
' Returns if this test is selected to be included during a test run.
'
' @return Returns True if the test will be run, False otherwise.
'
Public Property Get Selected() As Boolean: End Property

''
' Sets if this test will be included in the test run.
'
' @param RHS The new value of the selected flag.
'
Public Property Let Selected(ByVal RHS As Boolean): End Property

''
' Deselects this tests and all subtests.
'
Public Sub DeselectAll(): End Sub

''
' Selects this test and all subtests.
'
Public Sub SelectAll(): End Sub

''
' Returns if DoEvents will be called during test runs.
'
' @return Returns True if DoEvents is allowed, False otherwise.
'
Public Property Get AllowDoEvents() As Boolean: End Property

''
' Sets if DoEvents will be called during test runs.
'
' @param RHS The setting for DoEvents being called.
' @remarks DoEvents can be called during tests to allow for the Abort
' method to be called.
'
Public Property Let AllowDoEvents(ByVal RHS As Boolean): End Property

''
' Aborts the current set of tests.
'
Public Sub Abort(): End Sub

''
' Returns the full name of the test instance.
'
' @return Returns a fully qualified name for this test.
' @remarks The full name includes all parents for this test until
' no more parents exist.
'
Public Property Get FullName() As String: End Property

''
' Sorts the tests contained within and commands subtest to sort their tests.
'
' @param Comparer A custom comparer object used to perform custom sorting.
'
Public Sub Sort(Optional ByVal Comparer As ITestComparer): End Sub

''
' Returns the total number of tests under this test, including subtests.
'
' @return The total number of tests and subtests.
'
Public Property Get TotalTestCount() As Long: End Property

''
' Returns the filter to be applied to tests below this test level.
'
' @return Returns the current filter or Nothing.
'
Public Property Get Filter() As ITestFilter: End Property

''
' Sets the filter to be applied to tests below this test level.
'
' @param RHS The filter to be applied or Nothing.
'
Public Property Set Filter(ByVal RHS As ITestFilter): End Property

''
' Resets all tests back to the non-aborted stated.
'
Public Sub ClearAbort(): End Sub

