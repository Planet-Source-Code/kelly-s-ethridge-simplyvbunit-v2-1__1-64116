VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSuite"
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
'    Module: TestSuite
'

''
' Represents a collection of all test types to be executed as a single test.
'
' @remarks A <b>TestSuite</b> can contain <b>TestSuite</b>s, <b>TestCase</b>s,
' and <b>TestMethod</b>s to be executed as a single set of tests.
' <p>An object this is not one of these test types will be automatically wrapped
' in a <b>TestCase</b> to be executed as a single set of tests within the object.
' This allows a user to easily create a class of unit-tests and only need to pass
' in an instance of the class to have them executed within the <b>TestSuite</b>.</p>
'
' <pre>
' Dim Suite As New TestSuite
' Suite.Add New TestClass1
' Suite.Add New TestClass2
'
' Dim Result As ITestResult
' Set Result = Suite.Run
' </pre>
'
' @see Constructors
' @see ITest
' @see TestSuiteResult
' @see TestMethod
' @see TestCase
'
Option Explicit
Implements ITest
Implements ITestInternal



Private mTests          As New TestList
Private mName           As String
Private mRunner         As New TestRunner
Attribute mRunner.VB_VarHelpID = -1
Private mIsRunning      As Boolean
Private mParent         As Long
Private mSelected       As Boolean
Private mAllowDoEvents  As Boolean
Private mAbort          As Boolean
Private mFilter         As ITestFilter



''
' Performs a sort on the subtests contained with in the <b>TestSuite</b>.
'
' @param Comparer A custom comparer that can be used to sort the tests
' in a specific manner.
' @remarks All subtests are sorted that are contained in this suite. That
' includes tests contained in other test suites within this one.
' <p>If no Comparer is supplied, the <b>DefaultTestComparer</b> is used.</p>
'
Public Sub Sort(Optional ByVal Comparer As ITestComparer)
    Call mTests.Sort(Comparer)
    
    Dim Test As ITest
    For Each Test In mTests
        Call Test.Sort
    Next Test
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
' Returns if this test is currently running.
'
' @return Returns True if running, False otherwise.
'
Public Property Get IsRunning() As Boolean
    IsRunning = mRunner.IsRunning Or mIsRunning
End Property

''
' Adds tests to the suite to be run.
'
' @param Tests The tests to be added to the suite.
' @param Name The name of the test.
' @return Returns the test just added. This is helpful when adding a user object.
' A new <b>TestCase</b> is created for the object, and the <b>TestCase</b> is returned.
' @remarks A user object, <b>TestSuite</b>, <b>TestCase</b> or <b>TestMethod</b>
' can be added to a <b>TestSuite</b>. If a user object is added, then the object
' is wrapped by a <b>TestCase</b> object for testing.
'
Public Function Add(ByVal Tests As Object, Optional ByVal Name As String) As ITest
    If mIsRunning Then
        Call Err.Raise(ERR_INVALIDOPERATION, "TestSuite.Add", "Cannot add tests while the current TestSuite is running.")
    End If
    If Tests Is Nothing Then
        Call Err.Raise(5, "TestSuite.Add", "Cannot add Nothing to the test suite.")
    End If
    
    Dim Test As ITest
    If TypeOf Tests Is ITestInternal Then
        Set Test = Tests
    Else
        Set Test = Sim.NewTestCase(Tests, Name)
    End If
    
    If Len(Name) = 0 Then
        Name = Test.Name
    End If

    Set Test.Parent = Me
    Set Add = Test
    
    On Error GoTo errTrap
    Call mTests.Add(Test, Name)
    Exit Function
    
errTrap:
    Call Err.Raise(5, "TestSuite.Add", FormatString("A test with the name '{0}' already exists in the suite.", Name))
End Function

''
' Runs this suite of tests, returning the result.
'
' @param Listener A object that will be notified of events during the test execution.
' @return A result containing details about the execution of the test.
' @remarks All tests contained within the suite will have their <b>Run</b>
' method called (with the Listener passed in.)
' <p>The result is a <b>TestSuiteResult</b> and can be inspected for additional
' result from tests run within the suite.</p>
'
Public Function Run(Optional ByVal Listener As ITestListener) As ITestResult
    Set Run = mRunner.Run(Me, Listener, mSelected)
End Function

''
' Returns the name of this test.
'
' @return The name of this test.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Returns the number items in the suite at the first level.
'
' @return The number of level one items in the suite.
' @remarks This does not count subtests.
'
Public Property Get Count() As Long
    Count = mTests.Count
End Property

''
' Returns the number of tests contained in this suite.
'
' @return The total number of tests.
' @remarks Returns the total number of tests contained within
' this suite. All subtests are summed with their count aswell.
'
Public Property Get TotalTestCount() As Long
    Dim Test As ITest
    For Each Test In mTests
        TotalTestCount = TotalTestCount + Test.TotalTestCount
    Next Test
End Property

''
' Returns a test within the suite.
'
' @param Index The name or ordinal of the test.
' @return The test within the suite.
' @remarks The index is one-based.
'
Public Property Get Item(ByVal Index As Variant) As ITest
Attribute Item.VB_UserMemId = 0
    Set Item = mTests(Index)
End Property

''
' Used to enumerate the tests.
'
' @return A enumerator.
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
    On Error GoTo 0
    
    If Not Find Is Nothing Then
        Exit Function
    End If
    
    Dim Test As ITest
    For Each Test In mTests
        Set Find = Test.Find(Name)
        If Not Find Is Nothing Then
            Exit Function
        End If
    Next Test
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
Friend Sub Init(ByVal Name As String)
    mName = Name
End Sub

Friend Function InternalRun(ByVal ShouldRun As Boolean) As ITestResult
    Dim Result As New TestSuiteResult
    Call Result.Init(Me)
    
    If mAllowDoEvents Then
        DoEvents
    End If
    
    If mAbort Then
        Set InternalRun = Result
        Exit Function
    End If
    
    
    If mSelected And ShouldRun And Filters.Include(Me) Then
        mIsRunning = True
        
        ' We add our filter here, not before, so that we
        ' don't filter ourselves out.
        Call Filters.Add(mFilter)
        Call Listener.TestSuiteStarted(Me)
        
        Dim Ticks As Currency
        Ticks = GetTicks
        Call RunSubTests(Result, True)
        Call Result.SetTime(GetTime(Ticks))
        Call Result.SetExecuted
        Call Listener.TestSuiteCompleted(Result)
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
Private Sub RunSubTests(ByRef Result As TestSuiteResult, ByVal ShouldRun As Boolean)
    Dim i As Long
    For i = 1 To mTests.Count
        Dim Test As ITestInternal
        Set Test = mTests(i)
        Call Result.Add(Test.Run(ShouldRun))
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mName = "TestSuite"
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
    ITest_IsTestCase = False
End Property

Private Property Get ITest_IsTestMethod() As Boolean
    ITest_IsTestMethod = False
End Property

Private Property Get ITest_IsTestSuite() As Boolean
    ITest_IsTestSuite = True
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
    ITest_TotalTestCount = TotalTestCount
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestInternal Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestInternal_Run(ByVal ShouldRun As Boolean) As ITestResult
    Set ITestInternal_Run = InternalRun(ShouldRun)
End Function
