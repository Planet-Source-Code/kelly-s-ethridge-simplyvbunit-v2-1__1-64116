VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseResult"
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
'    Module: TestCaseResult
'

''
' Represents the results from a <b>TestCase</b> being run.
'
' @remarks A <b>TestCaseResult</b> contains all the test results
' for the tests called in a class.
'
' @see ITestResult
' @see TestCase
' @see TestSuiteResult
' @see TestResult
'
Option Explicit
Implements ITestResult

Private mResults        As New Collection
Private mIsIgnored      As Boolean
Private mMessage        As String
Private mDescription    As String
Private mExecuted       As Boolean
Private mTest           As ITest
Private mTime           As Currency
Private mParent         As Long
Private mIsError        As Boolean
Private mCause          As String



''
' Returns if the result is an error.
'
' @return Returns True if an error occured, False otherwise.
'
Public Property Get IsError() As Boolean
    IsError = mIsError
End Property

''
' Returns if there is an error in this result or any sub-results.
'
' @return Returns True if an error exists in this result or any sub-results, False otherwise.
' @remarks An error is not the same as a failure. An error is an unhandled
' error, a failure is caused by a failed assertion.
'
Public Property Get HasError() As Boolean
    If mIsError Then
        HasError = True
        Exit Property
    End If
    
    Dim Result As ITestResult
    For Each Result In mResults
        If Result.IsError Then
            HasError = True
            Exit Property
        End If
    Next Result
End Property

''
' Returns the cause of the error.
'
' @return The cause of the error.
'
Public Property Get Cause() As String
    Cause = mCause
End Property

''
' Returns if this test result currently has a parent.
'
' @return Returns True if there is currenlty a parent, False otherwise.
' @remarks It is possible for the parent to be released leaving
' only this test result. In that case False will be returned.
'
Public Property Get HasParent() As Boolean
    HasParent = (mParent <> 0)
End Property

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
Public Property Get Parent() As ITestResult
    If HasParent Then
        Dim Unk As IUnknown
        Call CopyMemory(Unk, mParent, 4)
        Set Parent = Unk
        Call ZeroMemory(Unk, 4)
    End If
End Property

''
' Sets the parent of this test result.
'
' @param RHS The new parent or Nothing.
' @remarks A weak reference to the parent is kept. This test result
' will not keep the parent alive. When the parent terminates,
' it should manually set this parent Nothing to ensure
' this test result does not try to use it.
'
Public Property Set Parent(ByVal RHS As ITestResult)
    mParent = ObjPtr(RHS)
End Property

''
' Returns a full path name for this test result.
'
' @return A full path name.
' @remarks A test result can be several level down within
' other test results. This returns a path from the toplevel
' test result to the current test result.
'
Public Property Get FullName() As String
    Dim Test As ITest
    Set Test = mTest
    Do While Not Test Is Nothing
        FullName = "\" & Test.Name & FullName
        Set Test = Test.Parent
    Loop
End Property

''
' Returns the time the test took to execute in milliseconds.
'
' @return Returns the execution time in milliseconds.
'
Public Property Get Time() As Currency
    Time = mTime
End Property

''
' Returns the test that generated this result.
'
' @return The test that generated this result.
' @remarks This type of result can only be generated by a <b>TestCase</b> test.
'
Public Property Get Test() As ITest
    Set Test = mTest
End Property

''
' Returns if this <b>TestCase</b> was executed.
'
' @return Returns True if executed, False otherwise.
' @remarks Even if the <b>TestCase</b> was executed, it is
' possible that none of the tests within it were executed.
'
Public Property Get Executed() As Boolean
    Executed = mExecuted
End Property

''
' Returns if the test was ignored.
'
' @return Returns True if ignored, False otherwise.
' @remarks A <b>TestCase</b> is ignored if <b>Assert.Ignore</b>
' is called in a <b>TestFixtureSetup</b> method. If it is called
' in a <b>Setup</b> method, then the test case is not considered ignored.
'
Public Property Get IsIgnored() As Boolean
    IsIgnored = mIsIgnored
End Property

''
' Returns if the test is a failure.
'
' @return Returns True if there is a failure, False otherwise.
' @remarks If any internal test failed, then this returns False.
'
Public Property Get IsFailure() As Boolean
    Dim Result As ITestResult
    For Each Result In mResults
        If Result.IsFailure Then
            IsFailure = True
            Exit Property
        End If
    Next Result
End Property

''
' Returns if the test succeeded.
'
' @return Returns True if there were no failures, errors or ignores, False otherwise.
' @remarks All internal tests must pass for this test to be a success.
'
Public Property Get IsSuccess() As Boolean
    If IsFailure Then Exit Property
    If IsIgnored Then Exit Property
    If IsError Then Exit Property
    
    IsSuccess = True
End Property

''
' Returns the number of tests run.
'
' @return The number of internal tests run.
' @remarks This sums up all the tests run by this test.
' <p>This does not include itself in the count.</p>
'
Public Property Get TestsRun() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        TestsRun = TestsRun + Result.TestsRun
    Next Result
End Property

''
' Returns the number of assertions called by all the test methods.
'
' @return The number of assertion calls.
'
Public Property Get AssertCount() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        AssertCount = AssertCount + Result.AssertCount
    Next Result
End Property

''
' Returns the number of failed tests in this test.
'
' @return The number of failed tests.
' @remarks This does not include itself in the count.
'
Public Property Get FailureCount() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        FailureCount = FailureCount + Result.FailureCount
    Next Result
End Property

''
' Returns the name of the test this result was generated from.
'
' @return The name of the generating test.
'
Public Property Get TestName() As String
    TestName = mTest.Name
End Property

''
' Returns a message about the test failure.
'
' @return The message about the failure.
'
Public Property Get Message() As String
    Message = mMessage
End Property

''
' Returns a description about the test failure.
'
' @return The description about the failure.
'
Public Property Get Description() As String
    Description = mDescription
End Property

''
' Returns a test result from within this test result
'
' @param Index The name or ordinal of the test result.
' @return The test result requested.
'
Public Property Get Item(ByVal Index As Variant) As ITestResult
Attribute Item.VB_UserMemId = 0
    Set Item = mResults(Index)
End Property

''
' Returns the number of test results in this test result.
'
' @return The number of test results.
'
Public Property Get Count() As Long
    Count = mResults.Count
End Property

''
' Returns an enumerator
'
' @return The enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mResults.[_NewEnum]
End Function

''
' Searches for a test result by name.
'
' @param Name The name of the resutl to find.
' @return Returns the first occurence of the name.
'
Public Function Find(ByVal Name As String) As ITestResult
    On Error Resume Next
    Set Find = mResults(Name)
    Call Err.Clear
End Function

''
' Returns a formatted string of the description and message.
'
' @return A formatted message.
'
Public Function ToString() As String
    ToString = mDescription & IIf(Len(mMessage) > 0, " - " & mMessage, "")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Test As ITest)
    Set mTest = Test
End Sub

Friend Sub Add(ByVal Result As ITestResult)
    Call mResults.Add(Result, Result.TestName)
End Sub

Friend Sub SetIgnore(ByVal Result As ContextResult)
    mMessage = Result.Message
    mDescription = Result.Description
    mCause = Result.Cause
    mIsIgnored = True
End Sub

Friend Sub SetContextError(ByVal Result As ContextResult)
    mMessage = Result.Message
    mDescription = Result.Description
    mCause = Result.Cause
    mIsError = True
End Sub

Friend Sub SetTime(ByVal RHS As Currency)
    mTime = RHS
End Sub

Friend Sub SetExecuted()
    mExecuted = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Dim Result As ITestResult
    For Each Result In mResults
        Set Result.Parent = Nothing
    Next Result
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestResult Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ITestResult_AssertCount() As Long
    ITestResult_AssertCount = AssertCount
End Property

Private Property Get ITestResult_Count() As Long
    ITestResult_Count = Count
End Property

Private Property Get ITestResult_Description() As String
    ITestResult_Description = Description
End Property

Private Property Get ITestResult_FailureCount() As Long
    ITestResult_FailureCount = FailureCount
End Property

Private Property Get ITestResult_Executed() As Boolean
    ITestResult_Executed = Executed
End Property

Private Function ITestResult_Find(ByVal Name As String) As ITestResult
    Set ITestResult_Find = Find(Name)
End Function

Private Property Get ITestResult_FullName() As String
    ITestResult_FullName = FullName
End Property

Private Property Get ITestResult_HasParent() As Boolean
    ITestResult_HasParent = HasParent
End Property

Private Property Get ITestResult_IsError() As Boolean
    ITestResult_IsError = IsError
End Property

Private Property Get ITestResult_IsFailure() As Boolean
    ITestResult_IsFailure = IsFailure
End Property

Private Property Get ITestResult_IsIgnored() As Boolean
    ITestResult_IsIgnored = IsIgnored
End Property

Private Property Get ITestResult_IsSuccess() As Boolean
    ITestResult_IsSuccess = IsSuccess
End Property

Private Property Get ITestResult_IsTestCaseResult() As Boolean
    ITestResult_IsTestCaseResult = True
End Property

Private Property Get ITestResult_IsTestResult() As Boolean
    ' do nothing
End Property

Private Property Get ITestResult_IsTestSuiteResult() As Boolean
    ' do nothing
End Property

Private Property Get ITestResult_Item(ByVal Index As Variant) As ITestResult
    Set ITestResult_Item = Item(Index)
End Property

Private Property Get ITestResult_Message() As String
    ITestResult_Message = Message
End Property

Private Function ITestResult_NewEnum() As stdole.IUnknown
    Set ITestResult_NewEnum = NewEnum
End Function

Private Property Set ITestResult_Parent(ByVal RHS As ITestResult)
    Set Parent = RHS
End Property

Private Property Get ITestResult_Parent() As ITestResult
    Set ITestResult_Parent = Parent
End Property

Private Property Get ITestResult_Test() As ITest
    Set ITestResult_Test = Test
End Property

Private Property Get ITestResult_TestName() As String
    ITestResult_TestName = TestName
End Property

Private Property Get ITestResult_TestsRun() As Long
    ITestResult_TestsRun = TestsRun
End Property

Private Property Get ITestResult_Time() As Currency
    ITestResult_Time = Time
End Property
