VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullTestListener"
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
'    Module: NullTestListener
'

''
' Represents an ITestListener that does no processing.
'
' @remarks A new ITestListener object would perform the same function,
' however, using a NullTestListener shows intention.
'
Option Explicit
Implements ITestListener


Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)

End Sub

Private Sub ITestListener_RunCompleted(ByVal Result As ITestResult)

End Sub

Private Sub ITestListener_RunStarted(ByVal Test As ITest)

End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As TestCaseResult)

End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As TestCase)

End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As TestResult)

End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestStarted(ByVal Test As TestMethod)

End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Result As TestSuiteResult)

End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suit As TestSuite)

End Sub
