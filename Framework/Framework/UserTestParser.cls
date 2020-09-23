VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserTestParser"
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
'    Module: TestCaseParser
'

''
' Parses a user supplied object to find the test methods to be run.
'
' If the object to inspect implements the ITestCase interface then
' the object is called through that interface to retrieve the names
' of the functions to be called during tests.
'
Option Explicit

Private mContext    As TestContext
Private mTestNames  As Collection



Friend Sub Parse(ByVal UserTests As Object)
    Set mContext = New TestContext
    Set mTestNames = New Collection
    
    If TypeOf UserTests Is ITestCase Then
        Call ManualParse(UserTests)
    Else
        Call AutoParse(UserTests)
    End If
End Sub

Friend Property Get Context() As TestContext
    Set Context = mContext
End Property

Friend Property Get TestNames() As Collection
    Set TestNames = mTestNames
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsRunnableTest(ByVal Member As MemberInfo) As Boolean
    If Member.InvokeKind <> INVOKE_FUNC Then Exit Function
    If Member.ReturnType <> VT_VOID Then Exit Function
    If Member.Parameters.Count > 0 Then Exit Function
    
    IsRunnableTest = True
End Function

Private Sub AutoParse(ByVal UserTests As Object)
    Dim Info As InterfaceInfo
    Dim Member As MemberInfo
    
    On Error GoTo errTrap
    Set Info = tli.InterfaceInfoFromObject(UserTests)
    For Each Member In Info.Members
        If IsRunnableTest(Member) Then
            Call AddMethod(UserTests, Member.Name)
        End If
    Next Member
    
errTrap:
End Sub

Private Sub ManualParse(ByVal UserTests As ITestCase)
    Dim Tests As New TestCollection
    Call UserTests.AddTests(Tests)
    
    Dim Name As Variant
    For Each Name In Tests.Names
        Call AddMethod(UserTests, Name)
    Next Name
End Sub

Private Sub AddMethod(ByVal UserTests As Object, ByVal Name As String)
    Select Case LCase$(Name)
        Case "setup":               Call mContext.SetSetup(UserTests)
        Case "teardown":            Call mContext.SetTearDown(UserTests)
        Case "testfixturesetup":    Call mContext.SetTestFixtureSetup(UserTests)
        Case "testfixtureteardown": Call mContext.SetTestFixtureTearDown(UserTests)
        Case Else:                  Call mTestNames.Add(Name)
    End Select
End Sub
