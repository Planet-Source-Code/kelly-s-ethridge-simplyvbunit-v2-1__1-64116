VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAreEqualDatesAsserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSpecializedEqualityAssertions
'
Option Explicit

Public Sub TestEqualDates()
    Assert.AreEqualDates #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:00 AM#, ToTheSecond, "Dates should be on same second."
    Assert.AreEqualDates #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:59 AM#, ToTheMinute, "Dates should be on same minute."
    Assert.AreEqualDates #1/1/2001 8:00:00 AM#, #1/1/2001 8:59:59 AM#, ToTheHour, "Dates should be on same hour."
    Assert.AreEqualDates #1/1/2001#, #1/1/2001 12:59:59 PM#, ToTheDay, "Dates should be on same day."
    Assert.AreEqualDates #1/1/2001#, #1/31/2001 12:59:59 PM#, ToTheMonth, "Dates should be on same month."
    Assert.AreEqualDates #1/1/2001#, #12/31/2001 12:59:59 PM#, ToTheYear, "Dates should be in same year."
    Assert.AreEqualDates #1/7/2001#, #1/13/2001 12:59:59 PM#, ToTheWeek, "Dates should be in same week."
    Assert.AreEqualDates #1/1/2001#, #3/31/2001 12:59:59 PM#, ToTheQuarter, "Dates should be in same quarter."
End Sub

Public Sub TestFailingToTheSecond()
    On Error GoTo errTrap
    Assert.AreEqualDates #1/1/2001 8:00:00 AM#, #1/1/2001 8:00:01 AM#, ToTheSecond
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheMinute()
    On Error GoTo errTrap
    Assert.AreEqualDates #1/1/2001 8:00:59 AM#, #1/1/2001 8:01:00 AM#, ToTheMinute
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheHour()
    On Error GoTo errTrap
    Assert.AreEqualDates #1/1/2001 8:59:59 AM#, #1/1/2001 9:00:00 AM#, ToTheHour
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheDay()
    On Error GoTo errTrap
    Assert.AreEqualDates #1/1/2001 12:59:59 PM#, #1/2/2001#, ToTheDay
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheMonth()
    On Error GoTo errTrap
    Assert.AreEqualDates #12/31/2001 12:59:59 PM#, #1/1/2002#, ToTheMonth
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheYear()
    On Error GoTo errTrap
    Assert.AreEqualDates #12/31/2001 12:59:59 PM#, #1/1/2002#, ToTheYear
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheWeek()
    On Error GoTo errTrap
    Assert.AreEqualDates #1/6/2001 12:59:59 PM#, #1/7/2001#, ToTheWeek
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

Public Sub TestFailingToTheQuarter()
    On Error GoTo errTrap
    Assert.AreEqualDates #3/31/2001 8:59:59 PM#, #4/1/2001#, ToTheQuarter
    On Error GoTo 0
    Assert.Fail "AreEqualDates should fail."
    
errTrap:
    
    Assert.AreEqual Assert.AssertCode, Err.Number, "Wrong error raised."
End Sub

