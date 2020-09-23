VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparerStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: ComparerStatic
'

''
' Provides comparers used for comparing values for equality.
'
Option Explicit

Private mObjectComparer As New ObjectComparer
Private mValueComparer  As New ValueComparer


''
' Returns a shared Object comparer.
'
' @return A shared object comparer.
' @remarks This comparer uses the "Is" method of comparing. If
' non-objects are passed into the compare method, False will be
' returned. No error is raised.
'
Public Property Get ObjectComparer() As IEqualityComparer
    Set ObjectComparer = mObjectComparer
End Property

''
' Returns a shared Value comparer.
'
' @return A shared value comparer.
' @remarks This comparer uses the "=" method for comparing. If
' objects are passed into the compare method, False will be
' returned. No error is raised.
'
Public Property Get ValueComparer() As IEqualityComparer
    Set ValueComparer = mValueComparer
End Property

''
' Returns a comparer based on whether an object compare is needed or not.
'
' @param IsObject Used to select which comparer is returned.
' @return A shared comparer.
'
Public Function GetComparer(ByVal IsObject As Boolean) As IEqualityComparer
    If IsObject Then
        Set GetComparer = ObjectComparer
    Else
        Set GetComparer = ValueComparer
    End If
End Function
