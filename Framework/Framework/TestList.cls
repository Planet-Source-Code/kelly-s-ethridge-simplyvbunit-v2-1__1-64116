VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestList"
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
'    Module: TestList
'

''
' This is a simple replacement of a Collection to allow for sorting the tests.
'
Option Explicit

Private Const DEF_CAPACITY As Long = 32


Private mTestCollection As New Collection
Private mTests()        As ITest
Private mCapacity       As Long
Private mComparer       As ITestComparer



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTestCollection.[_NewEnum]
End Function

' Needs to be public to be Default.
Public Property Get Item(ByVal Index As Variant) As ITest
Attribute Item.VB_UserMemId = 0
    If VarType(Index) = vbString Then
        Set Item = mTestCollection(Index)
    Else
        Set Item = mTests(CLng(Index))
    End If
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Add(ByRef Test As ITest, ByVal Name As String)
    Call mTestCollection.Add(Test, Name)
    Call EnsureCapacity(Count)
    Set mTests(Count) = Test
End Sub

Friend Property Get Count() As Long
    Count = mTestCollection.Count
End Property

Friend Sub Sort(ByVal Comparer As ITestComparer)
    If Comparer Is Nothing Then
        Set mComparer = TestComparer.Default
    Else
        Set mComparer = Comparer
    End If
    
    Call QuickSort(1, Count)
    Set mComparer = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub QuickSort(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As ITest
    Do While Left < Right
        i = Left: j = Right: Set x = mTests((i + j) \ 2)
        Do
            Do While mComparer.Compare(mTests(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Compare(mTests(j), x) > 0: j = j - 1: Loop
            
            If i > j Then Exit Do
            If i < j Then
                Dim t As ITest
                Set t = mTests(i)
                Set mTests(i) = mTests(j)
                Set mTests(j) = t
            End If
            
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call QuickSort(Left, j)
            Left = i
        Else
            If i < Right Then Call QuickSort(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    If RequiredCapacity <= mCapacity Then
        Exit Sub
    End If
    
    Dim NewCapacity As Long
    NewCapacity = mCapacity * 2
    
    If RequiredCapacity > NewCapacity Then
        NewCapacity = RequiredCapacity
    End If
    
    ReDim Preserve mTests(1 To NewCapacity)
    mCapacity = NewCapacity
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ReDim mTests(1 To DEF_CAPACITY)
    mCapacity = DEF_CAPACITY
End Sub
