VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCase"
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
'    Module: TestCase
'

''
' This class wraps a user test class.
'
' @remarks A new <b>TestCase</b> can be created using the <b>NewTestCase</b> method.
' <pre>
' Dim Test as TestCase
' Set Test = Sim.NewTestCase(New MyTests)
'
' Dim Result As ITestResult
' Set Result = Test.Run
' </pre>
' <p>A <b>TestCase</b> can be added to a <b>TestSuite</b> aswell.
'
' @see Constructors
' @see ITest
' @see ITestCase
' @see TestCaseResult
' @see TestSuite
'
Option Explicit
Implements ITest
Implements ITestInternal

Private mUserTests      As Object
Private mTests          As New TestList
Private mName           As String
Private mContext        As TestContext
Private mRunner         As New TestRunner
Attribute mRunner.VB_VarHelpID = -1
Private mIsRunning      As Boolean
Private mParent         As Long
Private mSelected       As Boolean
Private mAllowDoEvents  As Boolean
Private mAbort          As Boolean
Private mFilter         As ITestFilter



''
' Performs a sort on the subtests contains within this test.
'
' @param Comparer A custom comparer that can be used to sort
' the tests in a specified manner.
' @remarks As a <b>TestCase</b> this will sort the individual
' functions to be called during a run.
' <p>If no Comparer is supplied, the <b>DefaultTestComparer</b> is used.</p>
'
Public Sub Sort(Optional ByVal Comparer As ITestComparer)
    Call mTests.Sort(Comparer)
End Sub

''
' Returns a full path name for this test.
'
' @return A full path name.
' @remarks A test can be several level down within
' other tests. This returns a path from the toplevel
' test to the current test.
'
Public Property Get FullName() As String
    FullName = "\" & mName
    
    Dim Test As ITest
    Set Test = Parent
    Do While Not Test Is Nothing
        FullName = "\" & Test.Name & FullName
        Set Test = Test.Parent
    Loop
End Property

''
' Returns if DoEvents is called during the tests.
'
' @return Returns True if DoEvents is allowed, False otherwise.
' @remarks This is to help with calling the Abort method during
' tests. While tests are running they do not allow events, so Buttons
' cannot process. This gives everything a chance to process, but can
' cause the tests to take longer to run.
'
Public Property Get AllowDoEvents() As Boolean
    AllowDoEvents = mAllowDoEvents
End Property

''
' Sets the flag to allow DoEvents to be called while running a test.
'
' @param RHS The new flag value.
' @remarks This is to help with calling the Abort method during
' tests. While tests are running they do not allow events, so Buttons
' cannot process. This gives everything a chance to process, but can
' cause the tests to take longer to run.
'
Public Property Let AllowDoEvents(ByVal RHS As Boolean)
    mAllowDoEvents = RHS
    
    Dim Test As ITest
    For Each Test In mTests
        Test.AllowDoEvents = RHS
    Next Test
End Property

''
' Signals this test and all subtests that an abort has been requested.
'
Public Sub Abort()
    mAbort = True

    Dim Test As ITest
    For Each Test In mTests
        Call Test.Abort
    Next Test
End Sub

''
' Returns if the test will be run.
'
' @return Returns True if the test will be run, False otherwise.
' @remarks When the run method is invoked, the selected flag will be
' evaluated to determine if this test should be run or not. This is
' helpful when this test is part of a suite of tests that will all
' be run at the same time.
' <p>This does not change the Selected property for any subtests.</p>
'
Public Property Get Selected() As Boolean
    Selected = mSelected
End Property

''
' Sets if this test will be run or not.
'
' @param RHS Set to True to include the test when run, False otherwise.
' @remarks When the run method is invoked, the selected flag will be
' evaluated to determine if this test should be run or not. This is
' helpful when this test is part of a suite of tests that will all
' be run at the same time.
' <p>This does not change the Selected property for any subtests.</p>
'
Public Property Let Selected(ByVal RHS As Boolean)
    mSelected = RHS
End Property

''
' Deselects this test and all subtests to not be included when run.
'
' @remarks All subtests are iterated through and have their
' Selected property set to False.
'
Public Sub DeselectAll()
    Selected = False
    
    Dim Test As ITest
    For Each Test In mTests
        Call Test.DeselectAll
    Next Test
End Sub

''
' Selects this test and all subtests to be included when run.
'
' @remarks All subtests are iterated through and have their
' Selected property set to True.
'
Public Sub SelectAll()
    Selected = True
    
    Dim Test As ITest
    For Each Test In mTests
        Call Test.SelectAll
    Next Test
End Sub

''
' Returns the name of this test.
'
' @return The test name.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Returns if the test is currently running.
'
' @return Returns True if the test is running, False otherwise.
'
Public Property Get IsRunning() As Boolean
    IsRunning = mRunner.IsRunning Or mIsRunning
End Property

''
' Runs this test and any sub-tests.
'
' @param Listener A callback use to handle test events.
' @return The result of this test being run.
'
Public Function Run(Optional ByVal Listener As ITestListener) As ITestResult
    Set Run = mRunner.Run(Me, Listener, mSelected)
End Function

''
' Returns the number of tests within this test.
'
' @return The number of tests within this test.
'
Public Property Get Count() As Long
    Count = mTests.Count
End Property

''
' Returns the specified test within the <b>TestCase</b>.
'
' @param Index A test name or ordinal.
' @return The test within the <b>TestCase</b>.
' @remarks The index is one-based.
'
Public Property Get Item(ByVal Index As Variant) As ITest
Attribute Item.VB_UserMemId = 0
    Set Item = mTests(Index)
End Property

''
' Returns an enumerator.
'
' @return The enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTests.NewEnum
End Function

''
' Searches for the specified test.
'
' @param Name The name of the test to find.
' @return Returns the test if found or Nothing if not found.
' @remarks This will search all internal tests until it
' finds the first match.
'
Public Function Find(ByVal Name As String) As ITest
    On Error Resume Next
    Set Find = mTests(Name)
    Call Err.Clear
End Function

''
' Returns if this test currently has a parent.
'
' @return Returns True if there is currenlty a parent, False otherwise.
' @remarks It is possible for the parent to be released leaving
' only this test. In that case False will be returned.
'
Public Property Get HasParent() As Boolean
    HasParent = (mParent <> 0)
End Property

''
' Returns the parent test for this test.
'
' @return The parent test or Nothing.
' @remarks If the test is part of a <b>TestCase</b> or
' a <b>TestSuite</b> then that test will be returned.
' If there is no parent or the parent has released its
' reference to this test, then Nothing will be returned.
' <p>A weak reference to the parent is kept. This test will
' not keep the parent alive. If the parent terminates, it should
' manually set the parent to this test to Nothing.
'
Public Property Get Parent() As ITest
    If HasParent Then
        Dim Unk As IUnknown
        Call CopyMemory(Unk, mParent, 4)
        Set Parent = Unk
        Call ZeroMemory(Unk, 4)
    End If
End Property

''
' Sets the parent of this test.
'
' @param RHS The new parent or Nothing.
' @remarks A weak reference to the parent is kept. This test
' will not keep the parent alive. When the parent terminates,
' it should manually set this parent Nothing to ensure
' this test does not try to use it.
'
Public Property Set Parent(ByVal RHS As ITest)
    mParent = ObjPtr(RHS)
End Property

''
' Returns the current filtering object or Nothing.
'
' @return The current filter if one is set, or Nothing.
'
Public Property Get Filter() As ITestFilter
    Set Filter = mFilter
End Property

''
' Sets the filtering to be used to selectively include
' specific tests.
'
' @param RHS The filter to be used to filter tests.
' @remarks The filter is not applied to the specific test
' that the filter was set on. Intead, all subtests will be
' evaluated by the filter.
' <p>Any filters coming in from a parent will be combined with
' this filter and all will be used in subtest filtering.
' No tests that are not contained within this test will be
' evaluated using the filter from this test.</p>
'
Public Property Set Filter(ByVal RHS As ITestFilter)
    Set mFilter = RHS
End Property

''
' Clears the abort flag for the <b>TestCase</b> and sub-tests.
'
Public Sub ClearAbort()
    mAbort = False
    
    Dim Test As ITest
    For Each Test In mTests
        Call Test.ClearAbort
    Next Test
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal UserTests As Object, ByVal Name As String)
    If UserTests Is Nothing Then
        Call Err.Raise(5, "TestCase", "UserTests cannot be Nothing.")
    End If
    
    Set mUserTests = UserTests
    
    If Len(Name) = 0 Then
        mName = TypeName(UserTests)
    Else
        mName = Name
    End If
    
    Call InitTestMethods(UserTests)
End Sub

Friend Function InternalRun(ByVal ShouldRun As Boolean) As ITestResult
    Dim Result As New TestCaseResult
    Call Result.Init(Me)
    
    If mAllowDoEvents Then
        DoEvents
    End If
    
    If mAbort Then
        Set InternalRun = Result
        Exit Function
    End If
    
    
    Dim Res As New TestResult
    If mSelected And ShouldRun And Filters.Include(Me) Then
        mIsRunning = True
        
        Call Filters.Add(mFilter)
        Call Listener.TestCaseStarted(Me)
        
        Dim Ticks As Currency
        Ticks = GetTicks
        
        Dim ContextResult As ContextResult
        Set ContextResult = mContext.RunTestFixtureSetup(Me)
        
        If ContextResult.IsIgnored Then
            Call Result.SetIgnore(ContextResult)
        ElseIf ContextResult.IsSuccess Then
            ' Run all the tests. Each test will run
            ' the Setup and TearDown methods if applicable.
            Call RunSubTests(Result, True)
            
            ' After a successful fixture setup we must
            ' tear it down even if the test failed.
            Set ContextResult = mContext.RunTestFixtureTearDown(Me)
            If ContextResult.IsFailure Then
                Call Res.SetContextError(ContextResult)
                Call Result.Add(Res)
            End If
        Else
            Call Result.SetContextError(ContextResult)
        End If
        
        Call Result.SetTime(GetTime(Ticks))
        Call Result.SetExecuted
        Call Listener.TestCaseCompleted(Result)
        Call Filters.Remove(mFilter)
        mIsRunning = False
    Else
        Call RunSubTests(Result, False)
    End If
    
    Set InternalRun = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunSubTests(ByRef Result As TestCaseResult, ByVal ShouldRun As Boolean)
    Dim i As Long
    For i = 1 To mTests.Count
        Dim Test As TestMethod
        Set Test = mTests(i)
        Call Result.Add(Test.InternalRun(ShouldRun))
    Next i
End Sub

Private Sub InitTestMethods(ByVal UserTests As Object)
    Dim Parser As New UserTestParser
    Call Parser.Parse(UserTests)
    Set mContext = Parser.Context
    
    Dim TestName As Variant
    For Each TestName In Parser.TestNames
        Dim Test As TestMethod
        Set Test = Sim.NewTestMethod(UserTests, TestName, False)
        Set Test.Context = mContext
        Set Test.Parent = Me
        Call mTests.Add(Test, TestName)
    Next TestName
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mSelected = True
End Sub

Private Sub Class_Terminate()
    Dim Test As ITest
    For Each Test In mTests
        Set Test.Parent = Nothing
    Next Test
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITest Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITest_Abort()
    Call Abort
End Sub

Private Property Let ITest_AllowDoEvents(ByVal RHS As Boolean)
    AllowDoEvents = RHS
End Property

Private Property Get ITest_AllowDoEvents() As Boolean
    ITest_AllowDoEvents = AllowDoEvents
End Property

Private Sub ITest_ClearAbort()
    Call ClearAbort
End Sub

Private Property Get ITest_Count() As Long
    ITest_Count = Count
End Property

Private Sub ITest_DeselectAll()
    Call DeselectAll
End Sub

Private Property Set ITest_Filter(ByVal RHS As ITestFilter)
    Set Filter = RHS
End Property

Private Property Get ITest_Filter() As ITestFilter
    Set ITest_Filter = Filter
End Property

Private Function ITest_Find(ByVal Name As String) As ITest
    Set ITest_Find = Find(Name)
End Function

Private Property Get ITest_FullName() As String
    ITest_FullName = FullName
End Property

Private Property Get ITest_HasParent() As Boolean
    ITest_HasParent = HasParent
End Property

Private Property Get ITest_IsTestCase() As Boolean
    ITest_IsTestCase = True
End Property

Private Property Get ITest_IsTestMethod() As Boolean
    ITest_IsTestMethod = False
End Property

Private Property Get ITest_IsTestSuite() As Boolean
    ITest_IsTestSuite = False
End Property

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Function ITest_NewEnum() As stdole.IUnknown
    Set ITest_NewEnum = NewEnum
End Function

Private Property Set ITest_Parent(ByVal RHS As ITest)
    Set Parent = RHS
End Property

Private Property Get ITest_Parent() As ITest
    Set ITest_Parent = Parent
End Property

Private Function ITest_Run(Optional ByVal Listener As ITestListener) As ITestResult
    Set ITest_Run = Run(Listener)
End Function

Private Sub ITest_SelectAll()
    Call SelectAll
End Sub

Private Property Let ITest_Selected(ByVal RHS As Boolean)
    Selected = RHS
End Property

Private Property Get ITest_Selected() As Boolean
    ITest_Selected = Selected
End Property

Private Sub ITest_Sort(Optional ByVal Comparer As ITestComparer)
    Call Sort(Comparer)
End Sub

Private Property Get ITest_TotalTestCount() As Long
    ITest_TotalTestCount = Count
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestInternal Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestInternal_Run(ByVal ShouldRun As Boolean) As ITestResult
    Set ITestInternal_Run = InternalRun(ShouldRun)
End Function
