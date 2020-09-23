VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SettingHandler"
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
'    Module: SettingHandler
'

Option Explicit

Private WithEvents mTests As TreeView
Attribute mTests.VB_VarHelpID = -1
Private WithEvents mAutoRun As CheckBox
Attribute mAutoRun.VB_VarHelpID = -1
Private mHostName As String

Private mExpandedNodes As Collection



Friend Sub Init(ByVal TestList As TreeView, ByVal AutoRun As CheckBox)
    Set mTestList = TestList
    Set mAutoRun = AutoRun
End Sub

Private Sub SaveExpandedNodes()
    If mExpandedNodes.Count > 0 Then
        Dim Expanded() As String
        ReDim Expanded(0 To mExpandedNodes.Count - 1)
        
        Dim i As Long
        Dim Key As Variant
        For Each Key In mExpandedNodes
            Expanded(i) = Key
            i = i + 1
        Next Key
        
        Call SaveSetting("SimplyVBUnit2", mHostName, "ExpandedNodes", Join(Expanded, "|"))
    Else
        On Error Resume Next
        Call DeleteSetting("SimplyVBUnit2", mHostName)
    End If
End Sub

Private Sub CreateExpandedNodeCollection()
    Set mExpandedNodes = New Collection
    
    If mTests.Nodes.Count > 0 Then
'        Dim expand As String
        
        ' Only save expanded settings for child nodes
        ' if the root node is expanded. If the root node
        ' is not expanded, it will become expanded if a
        ' child node is set to expanded. We are assuming
        ' that by contracting the root node that we do
        ' not want to see any child nodes expanded again.
        '
        If mTests.Nodes("\TestSuite").Expanded Then
'            Dim Expanded() As String
'            ReDim Expanded(0 To tvwTests.Nodes.Count - 1)
            
            Dim i As Long
            Dim n As Node
            For Each n In tvwTests.Nodes
                If n.Expanded Then
                    'Expanded(i) = n.Key
                    'i = i + 1
                    Call mExpandedNodes.Add(n.Key)
                End If
            Next n
            
'            ReDim Preserve Expanded(0 To i - 1)
'            expand = Join(Expanded, "|")
        End If
        
'        Call SaveSetting(mHostName, "Settings", "Expanded", expand)
        
'    Else
'        On Error Resume Next
'        Call DeleteSetting(mHostName, "Settings", "Expanded")
    End If
End Sub


Private Sub mTestList_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub mTestList_Collapse(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Call mExpandedNodes.Remove(Node.Key)
End Sub

Private Sub mTestList_Expand(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Call mExpandedNodes.Add(Node.Key, Node.Key)
End Sub
