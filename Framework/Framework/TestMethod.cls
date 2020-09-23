VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMethod"
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
'    Module: TestMethod
'

''
' A <b>TestMethod</b> represents a single function within a class that is called
' when running unit tests.
'
' @remarks For a function in a class to become a <b>TestMethod</b>, it must be
' declared as a Public Sub without any arguments.
' <pre>
' Public Sub MyTestMethod()
' ...
' End Sub
' </pre>
' Friend and Private scoped Function and Property methods will be skipped.
' Public Sub methods with any arguments will also be skipped.
' <p>A <b>TestMethod</b> can be directly created by calling <b>NewTestMethod</b>.
' <pre>
' Dim Test As TestMethod
' Set Test = Sim.NewTestMethod(New MyTests, "MyTestMethodName")
'
' Dim Result As ITestResult
' Set Result = Test.Run
' </pre>
' A <b>TestMethod</b> can be added directly to a <b>TestSuite</b>.
'
' @see Constructors
' @see ITest
' @see TestResult
' @see TestSuite
'
Option Explicit
Implements ITest
Implements ITestInternal

Private mUserTestCase   As Object
Private mName           As String
Private mContext        As TestContext
Private mRunFullContext As Boolean
Private mRunner         As New TestRunner
Attribute mRunner.VB_VarHelpID = -1
Private mIsRunning      As Boolean
Private mParent         As Long
Private mSelected       As Boolean
Private mAllowDoEvents  As Boolean
Private mAbort          As Boolean



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
End Property

''
' Signals this test and all subtests that an abort has been requested.
'
Public Sub Abort()
    mAbort = True
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
' Runs the test.
'
' @param Listener A callback listener used to monitor the progress of the test.
' @return The results of the test.
' @remarks This function is not re-entrant, meaning an error will be raised
' if this function is called while the test is running.
'
Public Function Run(Optional ByVal Listener As ITestListener) As ITestResult
    If Me.HasParent Then
        If TypeOf Me.Parent Is TestCase Then
            mRunFullContext = True
            Set Run = mRunner.Run(Me, Listener, mSelected)
            mRunFullContext = False
            Exit Function
        End If
    End If
        
    Set Run = mRunner.Run(Me, Listener, mSelected)
End Function

''
' Returns the name of the test.
'
' @return The name of the test.
' @remarks This will be the name of the Sub called.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Returns if the test is currently running.
'
' @return Indication that the test is running.
'
Public Property Get IsRunning() As Boolean
    IsRunning = mRunner.IsRunning Or mIsRunning
End Property

''
' Returns the number of tests.
'
' @return This will always return 1 for a <b>TestMethod</b>.
'
Public Property Get Count() As Long
    Count = 1
End Property

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
' Searches for the specified test by name.
'
' @param Name The test being searched for.
' @return Returns the test if it is found, otherwise Nothing.
' @remarks Since this is a <b>TestMethod</b> there is
' only this instance of test to be searched, so this
' function simply compares the name to itself. If they
' are equal, then the function returns a reference to
' this instance of <b>TestMethod</b>.
'
Public Function Find(ByVal Name As String) As ITest
    If StrComp(Name, mName, vbTextCompare) = 0 Then
        Set Find = Me
    End If
End Function

''
' Clears the abort flag for the <b>TestCase</b> and sub-tests.
'
Public Sub ClearAbort()
    mAbort = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal UserTestCase As Object, ByVal Name As String, ByVal RunInContext As Boolean, ByVal Parent As ITest)
    If UserTestCase Is Nothing Then
        Call Err.Raise(5, "TestMethod", "UserTestCase cannot be Nothing.")
    End If
    If Len(Name) = 0 Then
        Err.Raise 5, "TestMethod", "Name cannot be empty."
    End If
    
    Set mUserTestCase = UserTestCase
    mName = Name
    mParent = ObjPtr(Parent)
    
    If RunInContext Then
        Call CreateContext
    Else
        Set mContext = TestContext.EmptyContext
    End If
End Sub

Friend Property Set Context(ByVal RHS As TestContext)
    Set mContext = RHS
End Property

Friend Function InternalRun(ByVal ShouldRun As Boolean) As ITestResult
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Me)
    
    If mAllowDoEvents Then
        DoEvents
    End If
    
    If mAbort Then
        Set InternalRun = Result
        Exit Function
    End If
    
    If Selected And ShouldRun And Filters.Include(Me) Then
        mIsRunning = True
        If mRunFullContext Then
            Call RunFullContext(Result)
        Else
            Call RunPartialContext(Result, True)
        End If
        mIsRunning = False
    End If
    
    Set InternalRun = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateContext()
    Dim Parser As New UserTestParser
    Call Parser.Parse(mUserTestCase)
    Set mContext = Parser.Context
    mRunFullContext = True
End Sub

Private Sub RunTest(ByVal Result As TestResult, ByVal CanRun As Boolean)
    Call Listener.TestStarted(Me)
    
    If CanRun Then
        Dim OldAssertCount As Long
        OldAssertCount = Assert.SetAssertCount(0)
    
        Dim Ticks As Currency
        Ticks = GetTicks
        On Error Resume Next
        
        ' This line is where it all happens. Pretty anticlimactic, huh?
        Call tli.InvokeHookSub(mUserTestCase, mName, INVOKE_FUNC)
        
        Select Case Err.Number
            Case ERR_ASSERT_PASS    ' hurray!
            Case ERR_ASSERT_FAIL:   Call Result.SetFail(Assert.Result)
            Case ERR_ASSERT_IGNORE: Call Result.SetIgnore(Assert.Result)
            Case Else:              Call Result.SetError(Err, mUserTestCase)
        End Select
        On Error GoTo 0
        
        Call Result.SetTime(GetTime(Ticks))
        Call Result.SetExecute
        Call Result.SetAssertCount(Assert.SetAssertCount(OldAssertCount))
    End If
    
    Call Listener.TestCompleted(Result)
End Sub

Private Sub RunFullContext(ByVal Result As TestResult)
    Dim SetupRan As Boolean
    SetupRan = HandleSetup(mContext.RunTestFixtureSetup(Me), Result)
    
    Call RunPartialContext(Result, SetupRan)
    
    If SetupRan Then
        Call HandleTearDown(mContext.RunTestFixtureTearDown(Me), Result)
    End If
End Sub

Private Sub RunPartialContext(ByVal Result As TestResult, ByVal CanRun As Boolean)
    Dim SetupRan As Boolean
    
    If CanRun Then
        SetupRan = HandleSetup(mContext.RunSetup(Me), Result)
    End If
    
    Call RunTest(Result, CanRun And SetupRan)
    
    If SetupRan Then
        Call HandleTearDown(mContext.RunTearDown(Me), Result)
    End If
End Sub

Private Function HandleSetup(ByVal SetupResult As ContextResult, ByVal TestResult As TestResult) As Boolean
    With SetupResult
        If .IsIgnored Then
            Call TestResult.SetIgnore(Assert.Result)
        ElseIf Not .IsSuccess Then
            Call TestResult.SetContextError(SetupResult)
        Else
            HandleSetup = True
        End If
    End With
End Function

Private Sub HandleTearDown(ByVal TearDownResult As ContextResult, ByVal TestResult As TestResult)
    If TearDownResult.IsFailure Then
        Call TestResult.SetContextError(TearDownResult)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mSelected = True
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
    Selected = False
End Sub

Private Property Set ITest_Filter(ByVal RHS As ITestFilter)
    ' do nothing
End Property

Private Property Get ITest_Filter() As ITestFilter
    ' do nothing
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
    ITest_IsTestMethod = True
End Property

Private Property Get ITest_IsTestSuite() As Boolean
    ITest_IsTestSuite = False
End Property

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Function ITest_NewEnum() As stdole.IUnknown
    Dim c As New Collection
    Set ITest_NewEnum = c.[_NewEnum]
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
    Selected = True
End Sub

Private Property Let ITest_Selected(ByVal RHS As Boolean)
    Selected = RHS
End Property

Private Property Get ITest_Selected() As Boolean
    ITest_Selected = Selected
End Property

Private Sub ITest_Sort(Optional ByVal Comparer As ITestComparer)
    ' I'm already sorted!
End Sub

Private Property Get ITest_TotalTestCount() As Long
    ITest_TotalTestCount = 1    ' just ourself
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestInternal Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestInternal_Run(ByVal ShouldRun As Boolean) As ITestResult
    Set ITestInternal_Run = InternalRun(ShouldRun)
End Function
