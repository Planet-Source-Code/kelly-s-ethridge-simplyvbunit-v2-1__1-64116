VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestComparer"
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
'    Module: ITestComparer
'

''
' An interface used to compare tests when sorting.
'
' @remarks A class can implement this interface to perform custom test sorting.
' <p>By default the <b>DefaultTestComparer</b> is used when sorting.</p>
'
Option Explicit

''
' The method called to compare two tests.
'
' @param x A test to be compared.
' @param y A test to be compared.
' @return Returns a value indicating the relationship when the tests are compared.
' @remarks When comparing the tests, the relationship is determined in the result as follows:<br><br>
' If x < y then return a negative value<br>
' If x > y then return a positive value<br>
' If x = y then return zero
'
Public Function Compare(ByVal x As ITest, ByVal y As ITest) As Long: End Function

