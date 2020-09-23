VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctions"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: PublicFunctions
'

''
' Provides a set of functions for interacting with the system.
'
' @remarks This class is defined as GlobalMultiUse to allow users to easily
' perform system wide tasks.
'
Option Explicit


''
' A user adds a new test to the GUI runner through this function.
'
' @param Test The new test to be added.
' @param Name A name for this set of tests.
' @remarks Normally a user created class containing the tests is passed
' in to be run. However, any <b>ITest</b> (<b>TestSuite</b>, <b>TestCase</b>,
' <b>TestMethod</b>) or <b>ITestCase</b> test can also be added.
'
Public Sub AddTest(ByVal Test As Object, Optional ByVal Name As String)
    Call UserEvents.OnAddTest(Test, Name)
End Sub

''
' Writes a line of text to the console output in the GUI.
'
' @param Text The text to be written to the console output.
'
Public Sub WriteLine(Optional ByVal Text As String)
    Call UserEvents.OnWriteLine(Text)
End Sub

''
' Adds an <b>ITestListener</b> object to a collection of listeners
' that will receive events during the running of the tests.
'
' @param Listener The listener to receive callback messages.
'
Public Sub AddListener(ByVal Listener As ITestListener)
    Call UserEvents.OnAddListener(Listener)
End Sub

''
' Adds an <b>ITestFilter</b> object to a collection of filters
' that will be used to determine which tests to be included
' during the test run.
'
' @param Filter The filter to be added.
' @remarks Filters are combined and used together to determine
' which tests will be included when running.
'
Public Sub AddFilter(ByVal Filter As ITestFilter)
    Call UserEvents.OnAddFilter(Filter)
End Sub

''
' Removes the filter from the collection of filters.
'
' @param Filter The filter to be removed.
'
Public Sub RemoveFilter(ByVal Filter As ITestFilter)
    Call UserEvents.OnRemoveFilter(Filter)
End Sub

