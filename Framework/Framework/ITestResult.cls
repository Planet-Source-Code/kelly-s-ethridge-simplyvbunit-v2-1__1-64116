VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestResult"
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
'    Module: ITestResult
'

''
' Represents a common interface shared amonst all of the test result classes.
'
' @see TestResult
' @see TestCaseResult
' @see TestSuiteResult
'
Option Explicit


''
' Returns if the test failed.
'
' @return Returns True if the test failed, False otherwise.
' @remarks A test can fail by failing an Assert method or by an unexpected error.
'
Public Property Get IsFailure() As Boolean: End Property

''
' Returns the number of tests executed during the execution of the
' test that generated this result.
'
' @return The number of tests run.
'
Public Property Get TestsRun() As Long: End Property

''
' Returns the number of assertions performed during the test.
'
' @return The number of assertions (including sub-test assertions).
'
Public Property Get AssertCount() As Long: End Property

''
' Returns if the test was successful.
'
' @return Returns True if the test was successful, False otherwise.
' @remarks If the test was IsIgnored then this will alway return False.
'
Public Property Get IsSuccess() As Boolean: End Property

''
' Returns the name of the test.
'
' @return The name of the test that was Executed.
'
Public Property Get TestName() As String: End Property

''
' Returns a message for a failed test.
'
' @return A message describing why the test failed.
' @remarks This is supplied as one of the argument in the
' Assert methods to be used if the test fails.
'
Public Property Get Message() As String: End Property

''
' Returns the reason the Assertion failed the test.
'
' @return The reason an Assertion decided to the test failed.
' @remarks The framework returns a message describing the
' expectation of the assertion and why it decided to fail the test.
'
Public Property Get Description() As String: End Property

''
' Returns if the test method was IsIgnored by calling <b>Assert.Ignore</b>.
'
' @return Returns True if the test was IsIgnored, False otherwise.
' @remarks If <b>IsIgnored</b> is True, then the following flags are set:
' <ul>IsFailure = True</ul>
' <ul>Executed = True</ul>
' <p>If the test failed, then it should be checked that it was IsIgnored
' to confirm the failure or IsIgnored state.</p>
'
Public Property Get IsIgnored() As Boolean: End Property

''
' Returns if this test has been Executed.
'
' @returns A value indicating if the test has already been run.
' @remarks If the test is ignored by calling <b>Assert.Ignore</b>
' in the test function, it is still flagged as being Executed.
' If <b>Assert.Ignore</b> was called in Setup or TestFixtureSetup
' then the test is not flagged as being executed.
'
Public Property Get Executed() As Boolean: End Property

''
' Returns the number of failures that occured.
'
' @return The number of failures that occured during this test execution.
'
Public Property Get FailureCount() As Long: End Property

''
' Returns the test this result was generated from.
'
' @return The test that generated this result.
'
Public Property Get Test() As ITest: End Property

''
' Returns the number of test results in this test result.
'
' @return The number of test results.
'
Public Property Get Count() As Long: End Property

''
' Returns a test result from within this test result
'
' @param Index The name or ordinal of the test result.
' @return The test result requested.
'
Public Property Get Item(ByVal Index As Variant) As ITestResult: End Property
Attribute Item.VB_UserMemId = 0

''
' Returns an enumerator
'
' @return The enumerator
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Searches for a test result by name.
'
' @param Name The name of the resutl to find.
' @return Returns the first occurence of the name.
'
Public Function Find(ByVal Name As String) As ITestResult: End Function

''
' Returns the time the test took to execute in milliseconds.
'
' @return Returns the execution time in milliseconds.
'
Public Property Get Time() As Currency: End Property

''
' Returns a full path name for this test result.
'
' @return A full path name.
' @remarks A test result can be several level down within
' other test results. This returns a path from the toplevel
' test result to the current test result.
'
Public Property Get FullName() As String: End Property

''
' Returns if this test result currently has a parent.
'
' @return Returns True if there is currenlty a parent, False otherwise.
' @remarks It is possible for the parent to be released leaving
' only this test result. In that case False will be returned.
'
Public Property Get HasParent() As Boolean: End Property

''
' Returns the parent test result for this test result
'
' @return The parent test result or Nothing.
' @remarks If the test result is part of a <b>TestCaseResult</b> or
' a <b>TestSuiteResult</b> then that test result will be returned.
' If there is no parent or the parent has released its
' reference to this test result, then Nothing will be returned.
' <p>A weak reference to the parent is kept. This test result will
' not keep the parent alive. If the parent terminates, it should
' manually set the parent to this test result to Nothing.
'
Public Property Get Parent() As ITestResult: End Property

''
' Sets the parent of this test result.
'
' @param RHS The new parent or Nothing.
' @remarks A weak reference to the parent is kept. This test result
' will not keep the parent alive. When the parent terminates,
' it should manually set this parent Nothing to ensure
' this test result does not try to use it.
'
Public Property Set Parent(ByVal RHS As ITestResult): End Property

''
' Returns if the result is an error.
'
' @return Returns True if an error occured, False otherwise.
'
Public Property Get IsError() As Boolean: End Property

''
' Returns if this instance of a test result is a <b>TestResult</b> class.
'
' @returns Returns True if this result instance is a <b>TestResult</b>, False otherwise.
'
Public Property Get IsTestResult() As Boolean: End Property

''
' Returns if this instance of a test result is a <b>TestCaseResult</b> class.
'
' @returns Returns True if this result instance is a <b>TestCaseResult</b>, False otherwise.
'
Public Property Get IsTestCaseResult() As Boolean: End Property

''
' Returns if this instance of a test result is a <b>TestSuiteResult</b> class.
'
' @returns Returns True if this result instance is a <b>TestSuiteResult</b>, False otherwise.
'
Public Property Get IsTestSuiteResult() As Boolean: End Property

