VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRunner"
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
'    Module: TestRunner
'

''
' Provides a common run routine used by all the test types.
'
Option Explicit


Private mIsRunning As Boolean



''
' Returns if the current test is running.
'
Friend Property Get IsRunning() As Boolean
    IsRunning = mIsRunning
End Property

''
' Executes a standard set of methods each test type (TestSuite, TestCase, TestMethod) all share.
'
' This is only called if the public Run method is called on a test. Subtests do not call this
' because all running is performed using the InternalRunner once the initial test is started.
'
Friend Function Run(ByVal Test As ITestInternal, ByVal oListener As ITestListener, ByVal ShouldRun As Boolean) As ITestResult
    If mIsRunning Then
        Call Err.Raise(ERR_INVALIDOPERATION, "Run", "Cannot call run while the current Test is running.")
    End If
    
    Call Iterations.BeginIteration(oListener)
    
    mIsRunning = True
    
    Call Listener.RunStarted(Test)
    Set Run = Test.Run(ShouldRun)
    Call Listener.RunCompleted(Run)
    
    mIsRunning = False
    
    Call Iterations.EndIteration
End Function
