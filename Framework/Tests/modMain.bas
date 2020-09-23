Attribute VB_Name = "modMain"
Option Explicit

Private Sub Main()
    TestCore
    TestUsingSimplyVBUnitLib
End Sub

Private Sub TestUsingSimplyVBUnitLib()
    Dim Suite As New TestSuite
    Suite.Add New TestConditionAssertions
    Suite.Add New TestEqualityAssertions
    Suite.Add New TestInEqualityAssertions
    Suite.Add New TestATestMethod
    Suite.Add New TestATestCase
    Suite.Add New TestATestSuite
    Suite.Add New TestAreEqualDatesAsserts
    Suite.Add New TestSpecialEqualityAsserts
    Suite.Add New TestIgnore
    Suite.Add New TestComparisonAssertions
    Suite.Add New TestSelectedTests
    Suite.Add New TestGetTestsByName
    Suite.Add New TestMultiCastListener
    Suite.Add New TestITestCase
    Suite.Add New TestDefaultTestComparer
    Suite.Add New TestNameFilter
    Suite.Add New TestFilter
    Suite.Add New TestMultiCastFilter
    Suite.Add New TestSayHear
    
    Suite.Run New SimpleListener
End Sub

Private Sub TestCore()
    Dim Test As New TestCore
    Test.Run
End Sub
