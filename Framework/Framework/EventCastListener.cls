VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventCastListener"
Attribute VB_GlobalNameSpace = False
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
'    Module: EventCastListner
'

''
' Converts <b>ITestListener</b> function callbacks into events.
'
' @remarks Use an <b>EventCastListener</b> to provide events instead
' of having to code callbacks by implementing the <b>ITestListener</b> interface.
' <p>An instance of this class should be delcared using <b>WithEvents</b>
' and passed in as a listener when a test is run. The callbacks will be
' redirected to events which can be responded to as needed.
'
' @see ITestListener
' @see MultiCastListener
'
Option Explicit
Implements ITestListener


''
' This is a message sent by using the <b>Assert.Say</b> assertion.
'
' @param Message The message sent.
' @param Args Any arguments that may be needed by the listener.
' @remarks This is a method of communicating specific information
' from an executing test to the listener. The listener may or may not
' understand what the message is about.
'
Public Event Hear(ByVal Message As String, ByRef Args() As Variant)

''
' This is called when all tests have been run.
'
' @param Result The is the final result for running the tests.
' @remarks This is called once after all tests have been completed.
'
Public Event RunCompleted(ByVal Result As ITestResult)

''
' This is called when the entire set of tests is about to begin.
'
' @param Test The first test (the once that Run was called on).
' @remarks This is called in preparation of all the tests to be run.
'
Public Event RunStarted(ByVal Test As ITest)

''
' This is called when a <b>TestCase</b> has been completed.
'
' @param Result The results of running the test case.
'
Public Event TestCaseCompleted(ByVal Result As TestCaseResult)

''
' This is called when a <b>TestCase</b> is starting.
'
' @param TestCase The <b>TestCase</b> about to begin running.
'
Public Event TestCaseStarted(ByVal TestCase As TestCase)

''
' This is called when a single test is complete.
'
' @param Result The result of the single test.
'
Public Event TestCompleted(ByVal Result As TestResult)

''
' This is called when a single test is about to begin.
'
' @param Test The test that is about to begin.
'
Public Event TestStarted(ByVal Test As TestMethod)

''
' This is called when a <b>TestSuite</b> has completed.
'
' @param Result The results of running the <b>TestCase</b>.
'
Public Event TestSuiteCompleted(ByVal Result As TestSuiteResult)

''
' This is called when a <b>TestSuite</b> is about to begin.
'
' @param Suite The <b>TestSuite</b> this is about to begin.
'
Public Event TestSuiteStarted(ByVal Suite As TestSuite)

''
' This is called when the <b>Setup</b> method is called.
'
' @param Test The test for which the setup is being called.
' @param HasContext Returns if there is a <b>Setup</b> method being called.
' @remarks This method is always called, even if no setup was defined in the testcase.
' This is to allow for reporting that no setup was included.
'
Public Event SetupStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)

''
' This is called when the <b>Setup</b> has completed.
'
' @param Result The result of running the context.
' @param HasContext Returns if there is a <b>Setup</b> method being called.
' @remarks This method is always called, even if no setup was defined in the testcase.
' This is to allow for reporting that no setup was included.
'
Public Event SetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

''
' This is called when the <b>Teardown</b> is called.
'
' @param Test The test for which the teardown is being called.
' @param HasContext Returns if there is a <b>Teardown</b> method being called.
' @remarks This method is always called, even if no teardown was defined in the testcase.
' This is to allow for reporting that no teardown was included.
'
Public Event TeardownStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)

''
' This is called when <b>Teardown</b> has completed for this test.
'
' @param Result The result of the teardown execution.
' @param HasContext Returns if there is a <b>Teardown</b> method being called.
' @remarks This method is always called, even if no teardown was defined in the testcase.
' This is to allow for reporting that no teardown was included.
'
Public Event TeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

''
' This is called when the <b>TestFixtureSetup</b> is being started.
'
' @param Test The <b>TestCase</b> or <b>TestMethod</b> this setup fixture is for.
' @param HasContext Returns if there is a <b>TestFixtureSetup</b> method.
' @remarks This method is always called, even if no TestFixtureSetup was defined in the testcase.
' This is to allow for reporting that no TestFixtureSetup was included.
'
Public Event TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

''
' This is called when the <b>TestFixtureSetup</b> has completed.
'
' @param Result The result of the test fixture execution.
' @param HasContext Returns if there is a <b>TestFixtureSetup</b> method declared.
' @remarks This method is always called, even if no TestFixtureSetup was defined in the testcase.
' This is to allow for reporting that no TestFixtureSetup was included.
'
Public Event TestFixtureSetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)

''
' This is called when the <b>TestFixtureTeardown</b> is being started.
'
' @param Test The <b>TestCase</b> or <b>TestMethod</b> this fixture is for.
' @param HasContext Returns if there is a <b>TestFixtureTeardown</b> declared.
' @remarks This method is always called, even if no TestFixtureTeardown was defined in the testcase.
' This is to allow for reporting that no TestFixtureTeardown was included.
'
Public Event TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

''
' This is called when the <b>TestFixtureTeardown</b> has completed.
'
' @param Result The result of the test fixture execution.
' @param HasContext Returns if there is a <b>TestFixtureTeardown</b> declared.
' @remarks This method is always called, even if no TestFixtureTeardown was defined in the testcase.
' This is to allow for reporting that no TestFixtureTeardown was included.
'
Public Event TestFixtureTeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestListener Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)
    RaiseEvent Hear(Message, Args)
End Sub

Private Sub ITestListener_RunCompleted(ByVal Result As ITestResult)
    RaiseEvent RunCompleted(Result)
End Sub

Private Sub ITestListener_RunStarted(ByVal Test As ITest)
    RaiseEvent RunStarted(Test)
End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)
    RaiseEvent SetupCompleted(Result, HasContext)
End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)
    RaiseEvent SetupStarted(Test, HasContext)
End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)
    RaiseEvent TeardownCompleted(Result, HasContext)
End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As TestMethod, ByVal HasContext As Boolean)
    RaiseEvent TeardownStarted(Test, HasContext)
End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As TestCaseResult)
    RaiseEvent TestCaseCompleted(Result)
End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As TestCase)
    RaiseEvent TestCaseStarted(TestCase)
End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As TestResult)
    RaiseEvent TestCompleted(Result)
End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)
    RaiseEvent TestFixtureSetupCompleted(Result, HasContext)
End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean)
    RaiseEvent TestFixtureSetupStarted(Test, HasContext)
End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As ContextResult, ByVal HasContext As Boolean)
    RaiseEvent TestFixtureTeardownCompleted(Result, HasContext)
End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasContext As Boolean)
    RaiseEvent TestFixtureTeardownStarted(Test, HasContext)
End Sub

Private Sub ITestListener_TestStarted(ByVal Test As TestMethod)
    RaiseEvent TestStarted(Test)
End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Result As TestSuiteResult)
    RaiseEvent TestSuiteCompleted(Result)
End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suite As TestSuite)
    RaiseEvent TestSuiteStarted(Suite)
End Sub
