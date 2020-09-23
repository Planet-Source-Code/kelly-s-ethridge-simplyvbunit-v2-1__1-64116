VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrentTestLabelListener"
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
'    Module: CurrentTestLabelListener
'

Option Explicit

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mLabel As Label



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Label As Label, ByVal Listener As EventCastListener)
    Set mLabel = Label
    Set mListener = Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    mLabel.Caption = ""
End Sub

Private Sub mListener_TestCaseStarted(ByVal TestCase As SimplyVBUnitLib.TestCase)
    mLabel.Caption = TestCase.FullName
End Sub

Private Sub mListener_TestStarted(ByVal Test As SimplyVBUnitLib.TestMethod)
    mLabel.Caption = Test.FullName
End Sub

Private Sub mListener_TestSuiteStarted(ByVal Suite As SimplyVBUnitLib.TestSuite)
    mLabel.Caption = Suite.FullName
End Sub
