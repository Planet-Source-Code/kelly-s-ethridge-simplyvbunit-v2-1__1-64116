VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilteredTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' All of these tests are filtered out by the system filter.
'
Option Explicit

' this will be run because we are using the default settings
' for the NameFilter filter. The default is to only perform
' filtering on test methods, not test cases or test suites.
'
' This function is part of a test case, so it is included.
Public Sub TestFixtureSetup()
    Assert.Say "I'm still going to run, though!"
End Sub

' these will not be run because of the filter.
'
' The filter is not case-sensitive.
Public Sub FILTEREDTestOne()
    Assert.Fail "I should be filtered out."
End Sub

Public Sub filteredTestTwo()
    Assert.Fail "I should be filtered out."
End Sub
