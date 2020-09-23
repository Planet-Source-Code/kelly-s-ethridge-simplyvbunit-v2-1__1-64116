VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticClasses"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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
'    Module: StaticClasses
'

''
' Provides access to commonly used methods.
'
Option Explicit


''
' Provides easy access to constructors in the library.
'
' @Return A class containing constructor functions.
' @remarks This is accessed by typing <b>Sim.</b> in the IDE
' followed by the constructor desired to execute.
'
Public Function Sim() As Constructors
    Set Sim = modStaticClasses.Sim
End Function


''
' Provides access to the shared <b>Assertions</b> object.
'
' @return The shared <b>Assertions</b> object.
' @remarks The entire system uses
Public Function Assert() As Assertions
    Set Assert = modStaticClasses.Assert
End Function

''
' Returns a Null listener used to prevent invalid
' listeners from being used.
'
' @return A null listener.
' @remarks When an callback is made, no action is performed.
'
Public Function NullTestListener() As NullTestListenerStatic
    Set NullTestListener = modStaticClasses.NullTestListener
End Function

''
' Provides access to the <b>CompareStatic</b> class.
'
' @return Returns a referece to <b>CompareStatic</b>.
'
Public Function Comparer() As ComparerStatic
    Set Comparer = modStaticClasses.Comparer
End Function

''
' Provides access to the <b>TestComparer</b> class.
'
' @return Returns a reference to <b>TestComparer</b>.
'
Public Function TestComparer() As TestComparerStatic
    Set TestComparer = modStaticClasses.TestComparer
End Function
