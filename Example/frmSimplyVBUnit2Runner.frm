VERSION 5.00
Object = "{BF02AA53-52CE-47D8-876F-0D0A78F085A7}#2.0#0"; "SimplyVBUnitUI.ocx"
Begin VB.Form frmSimplyVBUnit2Runner 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8685
   Begin SimplyVBUnitUI.SimplyVBUnitCtl SimplyVBUnitCtl1 
      Height          =   5895
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10398
   End
End
Attribute VB_Name = "frmSimplyVBUnit2Runner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' frmSimplyVBUnit2Runner V2.0
'
Option Explicit

' Namespaces Available:
'       Assert.*            ie. Assert.AreEqual Expected, Actual
'
' Public Functions Availabe:
'       AddTest <TestObject>
'       AddListener <ITestListener Object>
'       AddFilter <ITestFilter Object>
'       RemoveFilter <ITestFilter Object>
'       WriteLine "Message"
'
' Adding a testcase:
'   Use AddTest <object>
'
' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a Setup/Teardown method to be run before and after every test.
'   4. (Optionally) Add a TestFixtureSetup/TestFixtureTeardown method to be run at the
'      before the first test and after the last test.
'   5. Add public Subs of the tests you want run. No parameters.
'
'      Public Sub MyTest()
'          Assert.AreEqual a, b
'      End Sub
'
Private Sub Form_Load()
    ' Add tests here
    '
    ' AddTest <TestObject>
    
    ' I want all Assert.AreEqual string comparisons
    ' to be case insensitive, the default is vbBinaryCompare.
    Assert.StringCompareMethod = vbTextCompare
    
    ' Lets exclude tests that start with the name "Filter".
    Dim Filter As New NameFilter
    Filter.Pattern = "Filter*"
    Filter.Negate = True
    
    ' Add our new filter to the system.
    AddFilter Filter
    
    
    AddTest New CompilableTests
    AddTest New TestDifferentAsserts
    AddTest New RunMyOwnTests
    AddTest New FilteredTests
    
End Sub



Private Sub Form_Initialize()
    Caption = "SimplyVBUnit V2 - " & App.Title
    Call Me.SimplyVBUnitCtl1.Init(App.EXEName)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub


