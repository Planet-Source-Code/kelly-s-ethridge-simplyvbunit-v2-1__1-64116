VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl SimplyVBUnitCtl 
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ScaleHeight     =   6075
   ScaleWidth      =   6930
   Begin VB.CheckBox chkApplyToTestSuites 
      Caption         =   "Test Suites"
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Tag             =   "skip"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox chkApplyToTestCases 
      Caption         =   "Test Cases"
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Tag             =   "skip"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox chkApplyToTests 
      Caption         =   "Tests"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Tag             =   "skip"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtFilterPattern 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Tag             =   "skip"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox cboIncludeExclude 
      Height          =   315
      ItemData        =   "SimplyVBUnitCtl.ctx":0000
      Left            =   120
      List            =   "SimplyVBUnitCtl.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "skip"
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imglTreeView 
      Left            =   2400
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimplyVBUnitCtl.ctx":0020
            Key             =   "Passed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimplyVBUnitCtl.ctx":00E0
            Key             =   "NotRun"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimplyVBUnitCtl.ctx":019E
            Key             =   "Failed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SimplyVBUnitCtl.ctx":025F
            Key             =   "Ignored"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   3480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5175
      ScaleWidth      =   75
      TabIndex        =   9
      Tag             =   "skip"
      Top             =   0
      Width           =   75
      Begin VB.Image imgSplitterHandle 
         Height          =   300
         Left            =   600
         Picture         =   "SimplyVBUnitCtl.ctx":0321
         Top             =   2400
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.ListBox lstFailures 
      Appearance      =   0  'Flat
      Height          =   2700
      IntegralHeight  =   0   'False
      Left            =   3720
      TabIndex        =   10
      Top             =   2400
      Width           =   3015
   End
   Begin MSComctlLib.StatusBar statStatus 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Tag             =   "skip"
      Top             =   5685
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Status: Ready"
            TextSave        =   "Status: Ready"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "SuiteCount"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "CaseCount"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "TestCount"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Tests Run:"
            TextSave        =   "Tests Run:"
            Key             =   "TestsRun"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Failures:"
            TextSave        =   "Failures:"
            Key             =   "FailureCount"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "Time:"
            TextSave        =   "Time:"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Frame framControlBox 
      Height          =   1695
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Tag             =   "skip"
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkAutoRun 
         Appearance      =   0  'Flat
         Caption         =   "Automatically Run Tests"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "skip"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Tag             =   "skip"
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pbrTestRun 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Tag             =   "width"
         Top             =   1200
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblCurrentTest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Tag             =   "width"
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.TreeView tvwTests 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Tag             =   "skip"
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imglTreeView"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.ListBox lstOutput 
      Appearance      =   0  'Flat
      Height          =   2700
      IntegralHeight  =   0   'False
      Left            =   3720
      TabIndex        =   11
      Top             =   2400
      Width           =   3015
   End
   Begin MSComctlLib.TabStrip tabOutputs 
      Height          =   3135
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Errors and Failures"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Console Output"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblApplyFilterTo 
      Caption         =   "Apply to:"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Tag             =   "skip"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuExpandAll 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandTestSuite 
         Caption         =   "Expand TestSuites"
      End
      Begin VB.Menu mnuCollapseTestSuite 
         Caption         =   "Collapse TestSuites"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandTestCases 
         Caption         =   "Expand TestCases"
      End
      Begin VB.Menu mnuCollapsTestCases 
         Caption         =   "Collapse TestCases"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
      End
      Begin VB.Menu mnuRunAll 
         Caption         =   "Run All"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResult 
         Caption         =   "Result"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRealTimeUpdate 
         Caption         =   "Realtime Update"
      End
   End
End
Attribute VB_Name = "SimplyVBUnitCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: SimplyVBUnitCtl
'

Option Explicit

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const LB_SETHORIZONTALEXTENT As Long = &H194

Private Const SM_CYCAPTION      As Long = 4
Private Const SM_CYFRAME        As Long = 33
Private Const SM_CXFRAME        As Long = 32

Private Const MINIMUM_LEFT      As Long = 2500
Private Const MINIMUM_RIGHT     As Long = 3000

Private Const REG_APPNAME           As String = "SimplyVBUnitV2"
Private Const REG_AUTORUN           As String = "AutoRun"
Private Const REG_EXPANDEDNODES     As String = "ExpandedNodes"
Private Const REG_SPLITTERPOS       As String = "SplitterPos"
Private Const REG_SELECTEDTEST      As String = "SelectedTest"
Private Const REG_FORMLEFT          As String = "FormLeft"
Private Const REG_FORMTOP           As String = "FormTop"
Private Const REG_FORMWIDTH         As String = "FormWidth"
Private Const REG_FORMHEIGHT        As String = "FormHeight"
Private Const REG_TREEVIEWREDRAW    As String = "TreeViewRedraw"
Private Const REG_FILTERPATTERN     As String = "FilterPattern"
Private Const REG_FILTERNEGATE      As String = "FilterNegate"
Private Const REG_FILTERTESTS       As String = "FilterTests"
Private Const REG_FILTERTESTCASES   As String = "FilterTestCases"
Private Const REG_FILTERTESTSUITES  As String = "FilterTestSuites"


Private Const EXP_TESTSUITE     As Long = 2
Private Const EXP_TESTCASE      As Long = 1
Private Const EXP_ALL           As Long = 0

Private WithEvents mForm        As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mUserEvents  As UserEvents
Attribute mUserEvents.VB_VarHelpID = -1
Private mAnchor                 As Anchor
Private mHeightOffset           As Long
Private mWidthOffset            As Long
Private mTests                  As New Collection   ' this is so we can find a test by it's full name.
Private mHostName               As String
Private mRoot                   As TestSuite
Private mSplitting              As Boolean
Private mDX                     As Long
Private mOriginalLeft           As Long


' Listeners
Private WithEvents mListener    As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mProgressBarListener    As New ProgressBarListener
Private mFailureOutputListener  As New FailureOutputListener
Private mTreeViewListener       As New TreeViewListener
Private mStatusBarListener      As New StatusBarListener
Private mMultiListener          As New MultiCastListener
Private mLabelListener          As New CurrentTestLabelListener
Private mResultListener         As New TestResultCollection

Private mMultiFilter            As New MultiCastFilter
Private mFilter                 As New NameFilter




''
' Returns the width of control.
'
' @return Returns the width of the control.
'
Public Property Get Width() As Single
    Width = UserControl.Width
End Property

''
' Returns the height of the control.
'
' @return Returns the height of the control.
'
Public Property Get Height() As Single
    Height = UserControl.Height
End Property

''
' Called in the Form_Initialize event to start the GUI.
'
' @param HostEXEName The name of the application (App.EXEName)
'
Public Sub Init(ByVal HostEXEName As String)
    mHostName = HostEXEName
    Set mForm = UserControl.Parent
    mForm.Caption = "SimplyVBUnit v" & App.Major & "." & App.Minor
    
    Call SendMessage(lstFailures.hwnd, LB_SETHORIZONTALEXTENT, 1500, ByVal 0&)
    Call SendMessage(lstOutput.hwnd, LB_SETHORIZONTALEXTENT, 1500, ByVal 0&)
    Call LoadSettings
    Call DoAutoRun
End Sub

''
' Returns a collection that contains listeners.
'
' @return A collection of listeners.
' @remarks Use this to add new listeners to the test run. All listeners
' will receive the callback messages as the test run progresses.
' <p>Multiple listeners can be added.</p>
'
Public Property Get Listeners() As MultiCastListener
    Set Listeners = mMultiListener
End Property

''
' Returns a collection of filters.
'
' @return A collection of filters.
' @remarks Use this to add/remove filters from the test run.
' <p>Multiple filters can be added.</p>
'
Public Property Get Filters() As MultiCastFilter
    Set Filters = mMultiFilter
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayFilter()
    cboIncludeExclude.ListIndex = IIf(mFilter.Negate, 1, 0)
    txtFilterPattern.Text = mFilter.Pattern
    chkApplyToTests.Value = IIf(mFilter.FilterTestMethods, vbChecked, vbUnchecked)
    chkApplyToTestCases.Value = IIf(mFilter.FilterTestCases, vbChecked, vbUnchecked)
    chkApplyToTestSuites.Value = IIf(mFilter.FilterTestSuites, vbChecked, vbUnchecked)
End Sub

Private Sub InitFilter()
    Call mMultiFilter.Add(mFilter)
End Sub

Private Sub DoAutoRun()
    If Ambient.UserMode And (chkAutoRun.Value = vbChecked) Then
        Call RunSelectedTest
    End If
End Sub

Private Sub InitTreeView()
    With tvwTests
        Call .Nodes.Clear
        Call .Nodes.Add(, , mRoot.FullName, mRoot.Name, IMG_NOTRUN)
    End With
End Sub

Private Sub AddTestToTree(ByVal Parent As Node, ByVal Test As ITest)
    With tvwTests
        Dim Root As Node
        Set Root = .Nodes.Add(Parent, tvwChild, Test.FullName, Test.Name, IMG_NOTRUN)
        Call mTests.Add(Test, Test.FullName)
        
        Dim SubTest As ITest
        For Each SubTest In Test
            Call AddTestToTree(Root, SubTest)
        Next SubTest
        
        Root.Sorted = True
    End With
    
    Parent.Sorted = True
End Sub

Private Sub InitListeners()
    Call mLabelListener.Init(lblCurrentTest, mListener)
    Call mProgressBarListener.Init(pbrTestRun, mListener)
    Call mFailureOutputListener.Init(lstFailures, mListener)
    Call mTreeViewListener.Init(tvwTests, mListener, mnuRealTimeUpdate.Checked)
    Call mStatusBarListener.Init(statStatus, mListener)
    Call mResultListener.Init(mListener)
End Sub

Private Sub RunSelectedTest()
    Call Run(GetStartingTest)
End Sub

Private Sub RunAllTests()
    Call Run(mRoot)
End Sub

Private Sub Run(ByVal Test As ITest)
    Call lstOutput.Clear
    
    mRoot.Sort
    mRoot.AllowDoEvents = True
    Call mRoot.ClearAbort
    Call InitListeners
    
    With Test
        Set .Filter = Me.Filters
        Call .Run(Me.Listeners)
        Set .Filter = Nothing
    End With
End Sub

Private Function GetStartingTest() As ITest
    On Error GoTo errTrap
    Set GetStartingTest = mTests(tvwTests.SelectedItem.Key)
    Exit Function
    
errTrap:
    Set GetStartingTest = mRoot
End Function

Private Sub UpdateStatusBar()
    Dim SuiteCount  As Long
    Dim CaseCount   As Long
    Dim TestCount   As Long
    
    Call CountTestTypes(GetStartingTest, SuiteCount, CaseCount, TestCount)
    
    statStatus.Panels("SuiteCount").Text = "Test Suites: " & SuiteCount
    statStatus.Panels("CaseCount").Text = "Test Cases: " & CaseCount
    statStatus.Panels("TestCount").Text = "Tests: " & TestCount
End Sub

Private Sub CountTestTypes(ByVal Test As ITest, ByRef SuiteCount As Long, ByRef CaseCount As Long, ByRef TestCount As Long)
    If Test.IsTestCase Then
        CaseCount = CaseCount + 1
    ElseIf Test.IsTestSuite Then
        SuiteCount = SuiteCount + 1
    Else
        TestCount = TestCount + 1
    End If
    
    Dim t As ITest
    For Each t In Test
        Call CountTestTypes(t, SuiteCount, CaseCount, TestCount)
    Next t
End Sub

Private Sub SaveSettings()
    If Len(mHostName) = 0 Then Exit Sub
    
    Call SaveRegBoolean(REG_AUTORUN, chkAutoRun.Value = vbChecked)
    Call SaveExpandedNodes
    Call SaveRegSetting(REG_SPLITTERPOS, picSplitter.Left)
    Call SaveRegSetting(REG_SELECTEDTEST, GetStartingTest.FullName)
    Call SaveRegSetting(REG_FORMLEFT, mForm.Left)
    Call SaveRegSetting(REG_FORMTOP, mForm.Top)
    Call SaveRegSetting(REG_FORMWIDTH, mForm.Width)
    Call SaveRegSetting(REG_FORMHEIGHT, mForm.Height)
    Call SaveRegBoolean(REG_TREEVIEWREDRAW, mnuRealTimeUpdate.Checked)
    Call SaveRegSetting(REG_FILTERPATTERN, mFilter.Pattern)
    Call SaveRegBoolean(REG_FILTERNEGATE, mFilter.Negate)
    Call SaveRegBoolean(REG_FILTERTESTS, mFilter.FilterTestMethods)
    Call SaveRegBoolean(REG_FILTERTESTCASES, mFilter.FilterTestCases)
    Call SaveRegBoolean(REG_FILTERTESTSUITES, mFilter.FilterTestSuites)
End Sub

Private Sub LoadSettings()
    If Len(mHostName) = 0 Then Exit Sub
    
    mnuRealTimeUpdate.Checked = GetRegBoolean(REG_TREEVIEWREDRAW, True)
    chkAutoRun.Value = IIf(GetRegBoolean(REG_AUTORUN, False), vbChecked, vbUnchecked)
    
    Call LoadExpandedNodes
    Call LoadSelectedTest
    
    Dim Left As Long
    Dim Top As Long
    Dim Width As Long
    Dim Height As Long
    
    Left = GetRegSetting(REG_FORMLEFT, mForm.Left)
    Top = GetRegSetting(REG_FORMTOP, mForm.Top)
    Width = GetRegSetting(REG_FORMWIDTH, mForm.Width)
    Height = GetRegSetting(REG_FORMHEIGHT, mForm.Height)
    
    Call mForm.Move(Left, Top, Width, Height)
    
    Dim NewLeft As Long
    NewLeft = GetRegSetting(REG_SPLITTERPOS, picSplitter.Left)
    Call UpdateControlPositions(NewLeft - picSplitter.Left)
    picSplitter.Left = NewLeft
    
    mFilter.Negate = GetRegBoolean(REG_FILTERNEGATE, False)
    mFilter.Pattern = GetRegSetting(REG_FILTERPATTERN, "*")
    mFilter.FilterTestCases = GetRegBoolean(REG_FILTERTESTCASES, False)
    mFilter.FilterTestMethods = GetRegBoolean(REG_FILTERTESTS, True)
    mFilter.FilterTestSuites = GetRegBoolean(REG_FILTERTESTSUITES, False)
    Call DisplayFilter
End Sub

Private Sub LoadSelectedTest()
    On Error GoTo errTrap
    Dim n As Node
    Set n = tvwTests.Nodes(GetRegSetting(REG_SELECTEDTEST, mRoot.FullName))
    n.Selected = True

errTrap:
End Sub

Private Sub LoadExpandedNodes()
    Dim s As String
    s = GetRegSetting(REG_EXPANDEDNODES, mRoot.FullName)
    
    If Len(s) > 0 Then
        Dim Expanded() As String
        Expanded = Split(s, "|")
        
        ' node might not exist anymore.
        On Error Resume Next
        
        Dim i As Long
        For i = 0 To UBound(Expanded)
            tvwTests.Nodes(Expanded(i)).Expanded = True
        Next i
    End If
End Sub

Private Sub SaveExpandedNodes()
    If tvwTests.Nodes.Count > 0 Then
        Dim expand As String
        
        ' Only save expanded settings for child nodes
        ' if the root node is expanded. If the root node
        ' is not expanded, it will become expanded if a
        ' child node is set to expanded. We are assuming
        ' that by contracting the root node that we do
        ' not want to see any child nodes expanded again.
        '
        If tvwTests.Nodes(mRoot.FullName).Expanded Then
            Dim Expanded() As String
            ReDim Expanded(0 To tvwTests.Nodes.Count - 1)
            
            Dim i As Long
            Dim n As Node
            For Each n In tvwTests.Nodes
                If n.Expanded Then
                    Expanded(i) = n.Key
                    i = i + 1
                End If
            Next n
            
            ReDim Preserve Expanded(0 To i - 1)
            expand = Join(Expanded, "|")
        End If
        
        Call SaveRegSetting(REG_EXPANDEDNODES, expand)
        
    Else
        On Error Resume Next
        Call DeleteRegSetting(REG_EXPANDEDNODES)
    End If
End Sub

Private Sub InitAnchor(Optional ByVal Force As Boolean = False)
    If (mAnchor Is Nothing) Or Force Then
        Set mAnchor = New Anchor
        With mAnchor
            Call .Add(picSplitter, ToTop + ToBottom)
            Call .Add(tvwTests, ToTop + ToBottom)
            Call .Add(framControlBox, ToLeft + ToRight)
            Call .Add(pbrTestRun, ToLeft + ToRight)
            Call .Add(tabOutputs, ToAllSides)
            Call .Add(lstFailures, ToAllSides)
            Call .Add(lstOutput, ToAllSides)
            Call .Add(lblCurrentTest, ToLeft + ToRight)
            Call .Add(cboIncludeExclude, ToLeft + ToBottom)
            Call .Add(txtFilterPattern, ToLeft + ToBottom)
            Call .Add(lblApplyFilterTo, ToLeft + ToBottom)
            Call .Add(chkApplyToTests, ToLeft + ToBottom)
            Call .Add(chkApplyToTestCases, ToLeft + ToBottom)
            Call .Add(chkApplyToTestSuites, ToLeft + ToBottom)
        End With
    End If
End Sub

Private Sub UpdateControlPositions(ByVal ChangeAmount As Long)
    ' TreeView is special because it is on the other
    ' side of the Splitter bar.
    tvwTests.Width = tvwTests.Width + ChangeAmount
    
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        Select Case LCase$(c.Tag)
            Case "skip"
            Case "move" ' moves only, no changing width
                c.Left = c.Left + ChangeAmount
                
            Case "width"
                c.Width = c.Width - ChangeAmount
                
            Case Else   ' moves and changes width (narrows to the right)
                c.Left = c.Left + ChangeAmount
                c.Width = c.Width - ChangeAmount
        End Select
    Next c
End Sub

Private Sub SetExpand(ByVal Expanded As Boolean, ByVal ExpandType As Long)
    Call SuspendLayout(tvwTests.hwnd)
    
    Dim n As Node
    Dim skip As Boolean
    Dim Test As ITest
    For Each n In tvwTests.Nodes
        Select Case ExpandType
            Case EXP_TESTSUITE
                Set Test = mTests(n.Key)
                skip = Not Test.IsTestSuite
                
            Case EXP_TESTCASE
                Set Test = mTests(n.Key)
                skip = Not Test.IsTestCase
                
            Case Else
                skip = False
        
        End Select
        
        If Not skip Then
            n.Expanded = Expanded
        End If
    Next n
    
    Call ResumeLayout(tvwTests.hwnd)
End Sub

Private Sub SaveRegSetting(ByVal Key As String, ByVal Value As String)
    Call SaveSetting(REG_APPNAME, mHostName, Key, Value)
End Sub

Private Function GetRegSetting(ByVal Key As String, Optional ByVal Default As Variant) As String
    GetRegSetting = GetSetting(REG_APPNAME, mHostName, Key, Default)
End Function

Private Sub DeleteRegSetting(ByVal Key As String)
    On Error Resume Next
    Call DeleteSetting(REG_APPNAME, mHostName, Key)
End Sub

Private Function GetRegBoolean(ByVal Key As String, Optional ByVal Default As Boolean) As Boolean
    GetRegBoolean = CBool(GetRegSetting(Key, Default))
End Function

Private Sub SaveRegBoolean(ByVal Key As String, ByVal Value As Boolean)
    Call SaveRegSetting(Key, Value)
End Sub

Private Sub InitRoot()
    Set mRoot = Sim.NewTestSuite("All Tests")
    Call mTests.Add(mRoot, mRoot.FullName)
End Sub

Private Sub InitUserEvents()
    Set UserEvents = New UserEvents
    Set mUserEvents = UserEvents
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Control Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboIncludeExclude_Click()
    mFilter.Negate = (cboIncludeExclude.ListIndex = 1)
End Sub

Private Sub chkApplyToTestCases_Click()
    mFilter.FilterTestCases = (chkApplyToTestCases.Value = vbChecked)
End Sub

Private Sub chkApplyToTests_Click()
    mFilter.FilterTestMethods = (chkApplyToTests.Value = vbChecked)
End Sub

Private Sub chkApplyToTestSuites_Click()
    mFilter.FilterTestSuites = (chkApplyToTestSuites.Value = vbChecked)
End Sub

Private Sub cmdRun_Click()
    Call RunSelectedTest
End Sub

Private Sub cmdStop_Click()
    Call mRoot.Abort
End Sub

Private Sub mnuCollapseTestSuite_Click()
    Call SetExpand(False, EXP_TESTSUITE)
End Sub

Private Sub mnuCollapsTestCases_Click()
    Call SetExpand(False, EXP_TESTCASE)
End Sub

Private Sub mnuExpandTestCases_Click()
    Call SetExpand(True, EXP_TESTCASE)
End Sub

Private Sub mnuExpandTestSuite_Click()
    Call SetExpand(True, EXP_TESTSUITE)
End Sub

Private Sub mnuRealTimeUpdate_Click()
    mnuRealTimeUpdate.Checked = Not mnuRealTimeUpdate.Checked
End Sub

Private Sub mnuResult_Click()
    Call frmTestResult.ShowResult(mResultListener(tvwTests.SelectedItem.Key))
End Sub

Private Sub mnuRun_Click()
    Call RunSelectedTest
End Sub

Private Sub mnuRunAll_Click()
    Call RunAllTests
End Sub

Private Sub mUserEvents_AddFilter(ByVal Filter As SimplyVBUnitLib.ITestFilter)
    Call Me.Filters.Add(Filter)
End Sub

Private Sub mUserEvents_AddListener(ByVal Listener As SimplyVBUnitLib.ITestListener)
    Call Me.Listeners.Add(Listener)
End Sub

Private Sub mUserEvents_AddTest(ByVal Test As Object, ByVal Name As String)
    Call SuspendLayout(tvwTests.hwnd)
    Call AddTestToTree(tvwTests.Nodes(mRoot.FullName), mRoot.Add(Test, Name))
    Call ResumeLayout(tvwTests.hwnd)
    Call UpdateStatusBar
End Sub

Private Sub mUserEvents_RemoveFilter(ByVal Filter As SimplyVBUnitLib.ITestFilter)
    Call Me.Filters.Remove(Filter)
End Sub

Private Sub mUserEvents_WriteLine(ByVal Text As String)
    Call lstOutput.AddItem(Text)
End Sub

Private Sub mForm_Resize()
    Call UserControl.ParentControls(0).Controls(Extender.Name).Move(0, 0, mForm.Width - mWidthOffset, mForm.Height - mHeightOffset)
End Sub

Private Sub mnuCollapseAll_Click()
    Call SetExpand(False, EXP_ALL)
End Sub

Private Sub mnuExpandAll_Click()
    Call SetExpand(True, EXP_ALL)
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 And Button = vbLeftButton Then
        mSplitting = True
        mDX = X
        mOriginalLeft = picSplitter.Left
        picSplitter.BackColor = vbHighlight
    End If
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSplitting Then
        Dim NewX As Single
        NewX = picSplitter.Left + (X - mDX)
        If NewX < MINIMUM_LEFT Then
            NewX = MINIMUM_LEFT
        ElseIf NewX > (UserControl.Width - MINIMUM_RIGHT) Then
            NewX = UserControl.Width - MINIMUM_RIGHT
        End If
        
        picSplitter.Left = NewX
    End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSplitting And (Button = vbLeftButton) Then
        mSplitting = False
        picSplitter.BackColor = vbButtonFace
        Call UpdateControlPositions(picSplitter.Left - mOriginalLeft)
        Call InitAnchor(True)
        picSplitter.Refresh
    End If
End Sub

Private Sub picSplitter_Paint()
    Call picSplitter.Cls
    Call picSplitter.PaintPicture(imgSplitterHandle, 0, (picSplitter.Height - imgSplitterHandle.Height) / 2)
End Sub

Private Sub picSplitter_Resize()
    Call picSplitter.Refresh
End Sub

Private Sub tabOutputs_Click()
    lstFailures.Visible = tabOutputs.Tabs(1).Selected
    lstOutput.Visible = Not lstFailures.Visible
End Sub

Private Sub tvwTests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuTools
    End If
End Sub

Private Sub tvwTests_NodeClick(ByVal Node As MSComctlLib.Node)
    Call UpdateStatusBar
    If frmTestResult.Visible Then
        Call frmTestResult.ShowResult(mResultListener(Node.Key))
    End If
End Sub

Private Sub txtFilterPattern_LostFocus()
    mFilter.Pattern = txtFilterPattern.Text
End Sub

Private Sub UserControl_Hide()
    If Ambient.UserMode Then
        Set UserEvents = Nothing
        Call SaveSettings
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   UserControl Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Initialize()
    Call InitUserEvents
    Call InitRoot
    Call InitFilter
    
    Set mListener = New EventCastListener
    Call mMultiListener.Add(mListener)

    mHeightOffset = ScaleY(GetSystemMetrics(SM_CYFRAME) * 2 + GetSystemMetrics(SM_CYCAPTION), vbPixels, vbTwips)
    mWidthOffset = ScaleX(GetSystemMetrics(SM_CXFRAME) * 2, vbPixels, vbTwips)

    Call InitAnchor
    Call InitTreeView
End Sub

Private Sub UserControl_Resize()
    Call InitAnchor
    Call mAnchor.ReAnchor
End Sub
