VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FailureOutputListener"
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
'    Module: FailureOutputListener
'

Option Explicit

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mOutput As ListBox



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Output As ListBox, ByVal Listener As EventCastListener)
    Set mOutput = Output
    Set mListener = Listener
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFailure(ByVal Result As ITestResult)
    Dim s As String
    s = "<" & Result.Test.FullName & "> " & Result.Description & " - " & Result.Message
    Call mOutput.AddItem(s)
End Sub

Private Sub AddContextFailure(ByVal Result As ContextResult)
    Dim s As String
    s = "<" & Result.Test.FullName & "\" & Result.Name & "> " & Result.Description & " - " & Result.Message
    Call mOutput.AddItem(s)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_Hear(ByVal Message As String, Args() As Variant)
    Call mOutput.AddItem("---> Message: " & Message)
End Sub

Private Sub mListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    Call mOutput.Clear
End Sub

Private Sub mListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call AddContextFailure(Result)
    End If
End Sub

Private Sub mListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call AddContextFailure(Result)
    End If
End Sub

Private Sub mListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    If Result.IsFailure Or Result.IsError Or Result.IsIgnored Then
        Call AddFailure(Result)
    End If
End Sub

Private Sub mListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Or Result.IsIgnored Then
        Call AddContextFailure(Result)
    End If
End Sub

Private Sub mListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Or Result.IsIgnored Then
        Call AddContextFailure(Result)
    End If
End Sub
