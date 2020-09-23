VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgressBarListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ProgressBarListener
'

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const WM_USER           As Long = &H400
Private Const PBM_SETBARCOLOR   As Long = WM_USER + 9

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mProgressBar As ProgressBar
Private mCanChangeColor As Boolean




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Bar As ProgressBar, ByVal Listener As EventCastListener)
    Set mProgressBar = Bar
    Set mListener = Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetProgressBarColor(ByVal Color As OLE_COLOR)
    If mCanChangeColor Then
        Call SendMessage(mProgressBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal Color)
    End If
End Sub

Private Sub SetFailed()
    Call SetProgressBarColor(vbRed)
    mCanChangeColor = False
End Sub

Private Sub SetIgnored()
    Call SetProgressBarColor(vbYellow)
End Sub

Private Sub SetPassed()
    Call SetProgressBarColor(vbGreen)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunCompleted(ByVal Result As SimplyVBUnitLib.ITestResult)
    mProgressBar.Value = mProgressBar.Max
End Sub

Private Sub mListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    mCanChangeColor = True
    Call SetProgressBarColor(vbBlue)
    
    Dim Total As Long
    Total = Test.TotalTestCount
    If Total > 0 Then
        mProgressBar.Max = Total
    End If
    
    mProgressBar.Value = 0
End Sub

Private Sub mListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call SetFailed
    End If
End Sub

Private Sub mListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call SetFailed
    End If
End Sub

Private Sub mListener_TestCaseCompleted(ByVal Result As SimplyVBUnitLib.TestCaseResult)
    If Result.IsIgnored Then
        Call SetIgnored
        mProgressBar.Value = mProgressBar.Value + Result.Test.TotalTestCount
    End If
End Sub

Private Sub mListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    mProgressBar.Value = mProgressBar.Value + 1
    If Result.IsIgnored Then
        Call SetIgnored
    ElseIf Result.IsSuccess Then
        Call SetPassed
    Else
        Call SetFailed
    End If
End Sub

Private Sub mListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        mProgressBar.Value = mProgressBar.Value + Result.Test.TotalTestCount
        Call SetFailed
    End If
End Sub

Private Sub mListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call SetFailed
    End If
End Sub
