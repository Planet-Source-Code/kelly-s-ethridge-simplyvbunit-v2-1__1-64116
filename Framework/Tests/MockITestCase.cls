VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockITestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MockITestCase
'
Option Explicit
Implements ITestCase

Public RunSetup As Boolean
Public RunTeardown As Boolean
Public RunTestFixtureSetup As Boolean
Public RunTestFixtureTeardown As Boolean
Public RunTestOne As Boolean
Public RunTestTwo As Boolean
Public RunTestThree As Boolean
Public RunTestFour As Boolean
Public RunMissingTest As Boolean
Public RunInvalidTest As Boolean

Public SetupRunCount As Long
Public TeardownRunCount As Long
Public TestFixtureSetupRunCount As Long
Public TestFixtureTeardownRunCount As Long
Public TestOneRunCount As Long
Public TestTwoRunCount As Long
Public TestThreeRunCount As Long
Public TestFourRunCount As Long


Private Sub ITestCase_AddTests(ByVal Tests As SimplyVBUnitLib.TestCollection)
    If RunSetup Then Tests.Add "Setup"
    If RunTeardown Then Tests.Add "Teardown"
    If RunTestFixtureSetup Then Tests.Add "TestFixtureSetup"
    If RunTestFixtureTeardown Then Tests.Add "testfixtureteardown"
    If RunTestOne Then Tests.Add "TestOne"
    If RunTestTwo Then Tests.Add "TestTwo"
    If RunTestThree Then Tests.Add "TestThree"
    If RunTestFour Then Tests.Add "TestFour"
    If RunMissingTest Then Tests.Add "MissingTest"
    If RunInvalidTest Then Tests.Add "TestWithParameter"
End Sub

Public Sub TestWithParameter(ByVal i As Long)
    Assert.Fail "Should not be called."
End Sub

Public Sub Setup()
    Inc SetupRunCount
End Sub
Public Sub Teardown()
    Inc TeardownRunCount
End Sub
Public Sub TestFixtureSetup()
    Inc TestFixtureSetupRunCount
End Sub
Public Sub TestFixtureTeardown()
    Inc TestFixtureTeardownRunCount
End Sub
Public Sub TestOne()
    Inc TestOneRunCount
End Sub
Public Sub TestTwo()
    Inc TestTwoRunCount
End Sub
Public Sub TestThree()
    Inc TestThreeRunCount
End Sub
Public Sub TestFour()
    Inc TestFourRunCount
End Sub





Private Sub Inc(ByRef i As Long)
    i = i + 1
End Sub
