VERSION 5.00
Object = "{BF02AA53-52CE-47D8-876F-0D0A78F085A7}#1.0#0"; "SimplyVBUnitUI.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11700
   Begin SimplyVBUnitUI.SimplyVBUnitCtl SimplyVBUnitCtl1 
      Height          =   6255
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11033
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' frmMain
'
Option Explicit

Private Sub Form_Initialize()
    Me.SimplyVBUnitCtl1.Init App.EXEName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    
    AddTest New EmptyTestCast
    AddTest New FailingTestFixtureSetup
    AddTest New SimpleTestOne
    AddTest New SimpleTestTwo
    AddTest New IgnoreTestCase
    AddTest New FailingSetup
    

    Dim s As TestSuite
    Set s = Sim.NewTestSuite("A Suite")
    s.Add New SimpleTestOne
    AddTest s
    
End Sub

