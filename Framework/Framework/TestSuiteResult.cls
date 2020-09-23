VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSuiteResult"
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
'    Module: TestSuiteResult
'

''
' Represents the results of running a <b>TestSuite</b>.
'
' @remarks A <b>TestSuiteResult</b> contains all results run by
' sub-tests within the original <b>TestSuite</b>. Some of the
' functions rely on the sub-test results aswell.
'
' @see ITestResult
' @see TestSuite
' @see TestCaseResult
' @see TestResult
'
Option Explicit
Implements ITestResult

Private mResults    As New Collection
Private mTest       As ITest
Private mTime       As Currency
Private mExecuted   As Boolean
Private mParent     As Long



''
' Returns if the result is an error.
'
' @return Returns True if an error occured, False otherwise.
'
Public Property Get IsError() As Boolean
    Dim Result As ITestResult
    For Each Result In mResults
        If Result.IsError Then
            IsError = True
            Exit Property
        End If
    Next Result
End Property

''
' Returns if this this result has a parent.
'
' @return Returns True if this result has a parent, False otherwise.
'
Public Property Get HasParent() As Boolean
    HasParent = (mParent <> 0)
End Property

''
' Returns the parent of this test result.
'
' @return Returns this test result's parent if it has one, Nothing otherwise.
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
' Sets the parent for this test result.
'
' @param RHS The new parent of this test result.
' @remarks The test result only keeps a weak reference
' to the parent. The test result will not keep the parent alive.
'
Public Property Set Parent(ByVal RHS As ITestResult)
    mParent = ObjPtr(RHS)
End Property

''
' Returns a the full path name for this test result.
'
' @return The full path name.
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
' Returns if the test was executed.
'
' @return Returns True if the test was run, False otherwise.
'
Public Property Get Executed() As Boolean
    Executed = mExecuted
End Property

''
' Returns the test that generated this result.
'
' @return The test that generated this result.
'
Public Property Get Test() As ITest
    Set Test = mTest
End Property

''
' Returns the name of the test.
'
' @return The test name.
'
Public Property Get TestName() As String
    TestName = mTest.Name
End Property

''
' Returns the total number of tests run within the test.
'
' @return The number of tests run within the test.
'
Public Property Get TestsRun() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        TestsRun = TestsRun + Result.TestsRun
    Next Result
End Property

''
' Returns the number of failed tests within the test.
'
' @return The number of failed tests.
'
Public Property Get FailureCount() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        If Result.IsFailure Then
            FailureCount = FailureCount + Result.FailureCount
        End If
    Next Result
End Property

''
' Returns if the test is a failure.
'
' @return Returns True if the test failed, Falsed otherwise.
' @remarks Since this is a <b>TestSuiteResult</b> then all internal
' test results are examined for a failure. If any internal result
' fails, then this result is reported as a failure.
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
' @return If the test succeeded.
' @remarks For the test to succeed, all sub-tests must also succeed.
'
Public Property Get IsSuccess() As Boolean
    If IsFailure Then Exit Property
    If IsError Then Exit Property
    
    IsSuccess = True
End Property

''
' Returns the total number of assertions made while running the test.
'
' @return The total number of assertions made.
' @remarks All the assertions for the sub-tests are summed and returned.
'
Public Property Get AssertCount() As Long
    Dim Result As ITestResult
    For Each Result In mResults
        AssertCount = AssertCount + Result.AssertCount
    Next Result
End Property

''
' Returns a test result contained in this result.
'
' @param Index The name or ordinal of the test result.
' @return The test result requested.
'
Public Property Get Item(ByVal Index As Variant) As ITestResult
Attribute Item.VB_UserMemId = 0
    Set Item = mResults(Index)
End Property

''
' Returns the number of test results contained in this result.
'
' @return The number of test results.
' @remarks This does not include sub-test results. Only the
' first level of test results.
'
Public Property Get Count() As Long
    Count = mResults.Count
End Property

''
' Returns an enumerator
'
' @return The enumerator.
'
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
    On Error GoTo 0
    
    If Not Find Is Nothing Then
        Exit Function
    End If
    
    Dim Result As ITestResult
    For Each Result In mResults
        Set Find = Result.Find(Name)
        If Not Find Is Nothing Then
            Exit Function
        End If
    Next Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Add(ByVal Result As ITestResult)
    Call mResults.Add(Result, Result.TestName)
End Sub

Friend Sub Init(ByVal Test As ITest)
    Set mTest = Test
End Sub

Friend Sub SetExecuted()
    mExecuted = True
End Sub

Friend Sub SetTime(ByVal Time As Currency)
    mTime = Time
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
'   ITestResult_IsFailure
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ITestResult_AssertCount() As Long
    ITestResult_AssertCount = AssertCount
End Property

Private Property Get ITestResult_Count() As Long
    ITestResult_Count = Count
End Property

Private Property Get ITestResult_Description() As String
    ' do nothing
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
    
End Property

Private Property Get ITestResult_IsSuccess() As Boolean
    ITestResult_IsSuccess = IsSuccess
End Property

Private Property Get ITestResult_IsTestCaseResult() As Boolean
    ' do nothing
End Property

Private Property Get ITestResult_IsTestResult() As Boolean
    ' do nothing
End Property

Private Property Get ITestResult_IsTestSuiteResult() As Boolean
    ITestResult_IsTestSuiteResult = True
End Property

Private Property Get ITestResult_Item(ByVal Index As Variant) As ITestResult
    Set ITestResult_Item = Item(Index)
End Property

Private Property Get ITestResult_Message() As String
    ' do nothing
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
