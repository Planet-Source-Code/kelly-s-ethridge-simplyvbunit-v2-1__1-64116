VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestCase"
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
'    Module: ITestCase
'

''
' An interface used to allow private compiled objects to supply
' tests to be run by the framework.
'
' @remarks When a class is compiled and private, SimplyVBUnit cannot
' find the public routines to be used. An alternative is to manually
' specify the test names in the class to be executed. If a class
' implements this interface, it will be allowed to manually add test names
' to the test case that will be executed like normal.
' <p>If the class is not compiled or is not private, the test names collected
' through this interface will supersede any public methods in the class. This
' means if <b>Setup</b> is defined in the class, but not added using <b>Tests.Add</b>
' then it will not be called during test execution.</p>
' <p>A class that implements this interface can be added to a <b>TestSuite</b> or
' be used when creating a <b>TestCase</b> or <b>TestMethod</b>.
' <pre>
' Dim Suite As New TestSuite
' Suite.Add New MyITestCaseClass
' </pre>
' Or
' <pre>
' Dim TestCase As TestCase
' Set TestCase = NewTestCase(New MyITestCaseClass)
' </pre>
' Or
' <pre>
' Dim Test As TestMethod
' Set Test = NewTestMethod(New MyITestCaseClass, "MyTestMethod")
' </pre>
' <p>A class that implements <b>ITestCase</b> can be used in the normal manner
' as the previous test creation methods show.</p>
'
' @see Constructors
'
Option Explicit

''
' Callback routine used to allow test names to be added manually.
'
' @param Tests A collection to add test names to.
' @remarks When adding a test by name, the test must corrispond
' to the following method signature.
' <pre>
' Public Sub TestName()
' ...
' End Sub
' </pre>
' If context methods are to be included during the test execution, then
' they must be added to the list. For instance, if <b>Setup</b> is desired, then add
' "Setup" to the list. <b>Setup</b> will be called before the execution
' of each test in the list.
'
Public Sub AddTests(ByVal Tests As TestCollection): End Sub

