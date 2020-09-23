VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewListener"
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
'    Module: TreeViewListener
'

Option Explicit

Private mTreeView               As TreeView
Private WithEvents mListener    As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mRedraw                 As Boolean



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Tree As TreeView, ByVal Listener As EventCastListener, ByVal Redraw As Boolean)
    Set mTreeView = Tree
    Set mListener = Listener
    mRedraw = Redraw
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetStack(ByVal Node As Node, ByVal Image As String)
    Do While Not Node Is Nothing
        ' Fail has the highest priority, so leave it.
        If Node.Image = IMG_FAILED Then
            Exit Do
        End If
        
        Node.Image = Image
        Set Node = Node.Parent
    Loop
End Sub

Private Sub PassStack(ByVal Node As Node)
    Do While Not Node Is Nothing
        ' We don't hide any non-passing nodes.
        If Node.Image <> IMG_NOTRUN Then
            Exit Do
        End If

        Node.Image = IMG_PASSED
        Set Node = Node.Parent
    Loop
End Sub

' how horrible!
Private Sub IgnoreChildren(ByVal Node As Node)
    Call SetChildNodeImage(Node, IMG_IGNORED)
End Sub

Private Sub ClearNode(ByVal Node As Node)
    Call SetChildNodeImage(Node, IMG_NOTRUN)
End Sub

Private Sub SetChildNodeImage(ByVal Node As Node, ByRef Image As String)
    Do While Not Node Is Nothing
        If Node.Children > 0 Then
            Call SetChildNodeImage(Node.Child, Image)
        End If
        
        Node.Image = Image
        Set Node = Node.Next
    Loop
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   mListener Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mListener_RunCompleted(ByVal Result As SimplyVBUnitLib.ITestResult)
    If Not mRedraw Then
        Call ResumeLayout(mTreeView.hwnd)
    End If
End Sub

Private Sub mListener_RunStarted(ByVal Test As SimplyVBUnitLib.ITest)
    If Not mRedraw Then
        Call SuspendLayout(mTreeView.hwnd)
    End If
    
    Call ClearNode(mTreeView.Nodes(1))
End Sub

Private Sub mListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call SetStack(mTreeView.Nodes(Result.Test.FullName), IMG_FAILED)
    End If
End Sub

Private Sub mListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsFailure Then
        Call SetStack(mTreeView.Nodes(Result.Test.FullName), IMG_FAILED)
    End If
End Sub

Private Sub mListener_TestCaseCompleted(ByVal Result As SimplyVBUnitLib.TestCaseResult)
    Dim n As Node
    Set n = mTreeView.Nodes(Result.Test.FullName)
    
    If Result.IsIgnored Then
        Call SetStack(n, IMG_IGNORED)
        Call IgnoreChildren(n.Child)
    ElseIf Result.IsError Or Result.IsFailure Then
        Call SetStack(n, IMG_FAILED)
    End If
End Sub

Private Sub mListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    Dim n As Node
    Set n = mTreeView.Nodes(Result.Test.FullName)
    
    If Result.IsFailure Or Result.IsError Then
        Call SetStack(n, IMG_FAILED)
    ElseIf Result.IsIgnored Then
        Call SetStack(n, IMG_IGNORED)
    Else
        Call PassStack(n)
    End If
End Sub

Private Sub mListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)
    If Result.IsIgnored Then
        Call SetStack(mTreeView.Nodes(Result.Test.FullName), IMG_IGNORED)
    End If
End Sub

