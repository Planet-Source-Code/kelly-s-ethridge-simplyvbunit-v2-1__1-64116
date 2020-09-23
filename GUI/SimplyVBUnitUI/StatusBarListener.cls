VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarListener"
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
'    Module: StatusBarListener
'

Option Explicit

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mStatusBar As StatusBar



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Bar As StatusBar, ByVal Listener As EventCastListener)
    Set mStatusBar = Bar
    Set mListener = Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunCompleted(ByVal Result As SimplyVBUnitLib.ITestResult)
    mStatusBar.Panels("Status").Text = "Status: Finished"
    mStatusBar.Panels("FailureCount").Text = "Failures: " & Result.FailureCount
    mStatusBar.Panels("Time").Text = "Time: " & Result.Time & "ms"
    mStatusBar.Panels("TestsRun").Text = "Tests Run: " & Result.TestsRun
End Sub

Private Sub mListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    mStatusBar.Panels("Status").Text = "Status: Running"
End Sub

