VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IterationStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: IterationStack
'

Option Explicit

Private mIterations As New Collection


Friend Property Get Listener() As ITestListener
    If mIterations.Count > 0 Then
        Dim It As Iteration
        Set It = mIterations(1)
        Set Listener = It.Listener
    Else
        Set Listener = NullTestListener.NullListener
    End If
End Property

Friend Property Get Filters() As FilterList
    If mIterations.Count > 0 Then
        Dim It As Iteration
        Set It = mIterations(1)
        Set Filters = It.Filters
    End If
End Property

Friend Sub BeginIteration(ByVal Listener As ITestListener)
    Call EnsureHaveListener(Listener)
    
    Dim It As New Iteration
    Call It.Init(Listener)
    
    If mIterations.Count > 0 Then
        Call mIterations.Add(It, , 1)
    Else
        Call mIterations.Add(It)
    End If
End Sub

Friend Sub EndIteration()
    Call mIterations.Remove(1)
End Sub
