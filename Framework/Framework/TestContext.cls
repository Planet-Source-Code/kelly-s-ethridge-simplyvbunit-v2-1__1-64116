VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestContext"
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
'    Module: TestContext
'

''
' This is used to call any setup/teardown methods needed.
'
' This is filled by the UsertTestParser.
'
Option Explicit

Private mSetupMethod                  As ContextMethod
Private mTearDownMethod               As ContextMethod
Private mTestFixtureSetupMethod       As ContextMethod
Private mTestFixtureTearDownMethod    As ContextMethod



Friend Sub SetSetup(ByVal UserTests As Object)
    Set mSetupMethod = NewContextMethod(UserTests, "Setup")
End Sub

Friend Sub SetTearDown(ByVal UserTests As Object)
    Set mTearDownMethod = NewContextMethod(UserTests, "TearDown")
End Sub

Friend Sub SetTestFixtureSetup(ByVal UserTests As Object)
    Set mTestFixtureSetupMethod = NewContextMethod(UserTests, "TestFixtureSetup")
End Sub

Friend Sub SetTestFixtureTearDown(ByVal UserTests As Object)
    Set mTestFixtureTearDownMethod = NewContextMethod(UserTests, "TestFixtureTearDown")
End Sub

Friend Function RunSetup(ByVal Test As TestMethod) As ContextResult
    Call Listener.SetupStarted(Test, Not mSetupMethod Is Nothing)
    Set RunSetup = RunMethod(mSetupMethod, Test)
    Call Listener.SetupCompleted(RunSetup, Not mSetupMethod Is Nothing)
End Function

Friend Function RunTearDown(ByVal Test As TestMethod) As ContextResult
    Call Listener.TeardownStarted(Test, Not mTearDownMethod Is Nothing)
    Set RunTearDown = RunMethod(mTearDownMethod, Test)
    Call Listener.TeardownCompleted(RunTearDown, Not mTearDownMethod Is Nothing)
End Function

Friend Function RunTestFixtureSetup(ByVal Test As ITest) As ContextResult
    Call Listener.TestFixtureSetupStarted(Test, Not mTestFixtureSetupMethod Is Nothing)
    Set RunTestFixtureSetup = RunMethod(mTestFixtureSetupMethod, Test)
    Call Listener.TestFixtureSetupCompleted(RunTestFixtureSetup, Not mTestFixtureSetupMethod Is Nothing)
End Function

Friend Function RunTestFixtureTearDown(ByVal Test As ITest) As ContextResult
    Call Listener.TestFixtureTeardownStarted(Test, Not mTestFixtureTearDownMethod Is Nothing)
    Set RunTestFixtureTearDown = RunMethod(mTestFixtureTearDownMethod, Test)
    Call Listener.TestFixtureTeardownCompleted(RunTestFixtureTearDown, Not mTestFixtureTearDownMethod Is Nothing)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunMethod(ByVal Method As ContextMethod, ByVal Test As ITest) As ContextResult
    If Method Is Nothing Then
        Set RunMethod = TestContext.EmptyResult(Test)
    Else
        Set RunMethod = Method.Run(Test)
    End If
End Function

Private Function NewContextMethod(ByVal UserTests As Object, ByVal Name As String) As ContextMethod
    Set NewContextMethod = New ContextMethod
    Call NewContextMethod.Init(UserTests, Name)
End Function
