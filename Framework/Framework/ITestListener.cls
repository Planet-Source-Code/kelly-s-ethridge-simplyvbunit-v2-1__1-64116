VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestListener"
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
'    Module: ITestListener
'

''
' Provides an interface to allow a class to receive callback events during test execution.
'
' @remarks By implementing this interface, callbacks can be received to allow for
' custom reporting during test execution.
'
' @see MultiCastListener
' @see EventCastListener
'
Option Explicit

''
' This is called when a <b>TestMethod</b> is being started.
'
' @param Test The test that is being started.
'
Public Sub TestStarted(ByVal Test As TestMethod): End Sub

''
' This is called when a <b>TestMethod</b> has been completed.
'
' @param Result The result of the test just executed.
'
Public Sub TestCompleted(ByVal Result As TestResult): End Sub

''
' This is called when a <b>TestCase</b> is being started.
'
' @param TestCase The test case being executed.
'
Public Sub TestCaseStarted(ByVal TestCase As TestCase): End Sub

''
' This is called when a <b>TestCase</b> has been completed.
'
' @param Result This is the result of the <b>TestCase</b> execution.
'
Public Sub TestCaseCompleted(ByVal Result As TestCaseResult): End Sub

''
' This is called when a <b>TestSuite</b> is being started.
'
' @param Suite The test beint started.
'
Public Sub TestSuiteStarted(ByVal Suite As TestSuite): End Sub

''
' This is called when a <b>TestSuite</b> has completed executing.
'
' @param Result The result of the <b>TestSuite</b> execution.
'
Public Sub TestSuiteCompleted(ByVal Result As TestSuiteResult): End Sub

''
' This is called when <b>Run</b> is called on a test.
'
' @param Test The test that <b>Run</b> is called on.
' @remarks This is the test the Listener was passed into.
'
Public Sub RunStarted(ByVal Test As ITest): End Sub

''
' This is called when the full run has completed.
'
' @param Result The result of the total execution of the run.
'
Public Sub RunCompleted(ByVal Result As ITestResult): End Sub

''
' This is called when the <b>Assert.Say</b> method is called.
'
' @param Message A text message sent to the listener.
' @param Args Additional args that may be supplied to the listener.
'
Public Sub Hear(ByVal Message As String, ByRef Args() As Variant): End Sub

''
' This is called when the <b>Setup</b> method is called.
'
' @param Test The test for which the setup is being called.
' @param HasContext Returns if there is a <b>Setup</b> method being called.
' @remarks This method is always called, even if no setup was defined in the testcase.
' This is to allow for reporting that no setup was included.
'
Public Sub SetupStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean): End Sub

''
' This is called when the <b>Setup</b> has completed.
'
' @param Result The result of running the context.
' @param HasContext Returns if there is a <b>Setup</b> method being called.
' @remarks This method is always called, even if no setup was defined in the testcase.
' This is to allow for reporting that no setup was included.
'
Public Sub SetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean): End Sub

''
' This is called when the <b>Teardown</b> is called.
'
' @param Test The test for which the teardown is being called.
' @param HasContext Returns if there is a <b>Teardown</b> method being called.
' @remarks This method is always called, even if no teardown was defined in the testcase.
' This is to allow for reporting that no teardown was included.
'
Public Sub TeardownStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean): End Sub

''
' This is called when <b>Teardown</b> has completed for this test.
'
' @param Result The result of the teardown execution.
' @param HasContext Returns if there is a <b>Teardown</b> method being called.
' @remarks This method is always called, even if no teardown was defined in the testcase.
' This is to allow for reporting that no teardown was included.
'
Public Sub TeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean): End Sub

''
' This is called when the <b>TestFixtureSetup</b> is being started.
'
' @param Test The <b>TestCase</b> or <b>TestMethod</b> this setup fixture is for.
' @param HasContext Returns if there is a <b>TestFixtureSetup</b> method.
' @remarks This method is always called, even if no TestFixtureSetup was defined in the testcase.
' This is to allow for reporting that no TestFixtureSetup was included.
'
Public Sub TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean): End Sub

''
' This is called when the <b>TestFixtureSetup</b> has completed.
'
' @param Result The result of the test fixture execution.
' @param HasContext Returns if there is a <b>TestFixtureSetup</b> method declared.
' @remarks This method is always called, even if no TestFixtureSetup was defined in the testcase.
' This is to allow for reporting that no TestFixtureSetup was included.
'
Public Sub TestFixtureSetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean): End Sub

''
' This is called when the <b>TestFixtureTeardown</b> is being started.
'
' @param Test The <b>TestCase</b> or <b>TestMethod</b> this fixture is for.
' @param HasSetup Returns if there is a <b>TestFixtureTeardown</b> declared.
' @remarks This method is always called, even if no TestFixtureTeardown was defined in the testcase.
' This is to allow for reporting that no TestFixtureTeardown was included.
'
Public Sub TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasSetup As Boolean): End Sub

''
' This is called when the <b>TestFixtureTeardown</b> has completed.
'
' @param Result The result of the test fixture execution.
' @param HasContext Returns if there is a <b>TestFixtureTeardown</b> declared.
' @remarks This method is always called, even if no TestFixtureTeardown was defined in the testcase.
' This is to allow for reporting that no TestFixtureTeardown was included.
'
Public Sub TestFixtureTeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean): End Sub

