VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCollection"
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
'    Module: TestCollection
'

''
' Provides a container used to hold test names provided by a user
' during the <b>ITestCase_AddTests</b> callback method.
'
' @remarks This is to allow an object that can't be inspected by
' the framework to be allowed to provide the test names needed to be run.
' <p>When adding a test by name, the test must corrispond
' to the following method signature.
' <pre>
' Public Sub TestName()
' ...
' End Sub
' </pre></p>
' <p><b>Setup/Teardown</b> and <b>TestFixtureSetup/TestFixtureTeardown</b> are
' still called in the normal manner if they are included in the the test names.</p>
'
Option Explicit

Private mTestNames As New Collection



''
' Adds the name of a test to the list of tests to be run.
'
' @param TestName The name to be included when running the tests.
' @remarks When adding a test by name, the test must corrispond
' to the following method signature.
' <pre>
' Public Sub TestName()
' ...
' End Sub
' </pre>
'
Public Sub Add(ByVal TestName As String)
    Call mTestNames.Add(TestName, TestName)
End Sub

''
' Returns the number of names added so far.
'
Public Property Get Count() As Long
    Count = mTestNames.Count
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get Names() As Collection
    Set Names = mTestNames
End Property
