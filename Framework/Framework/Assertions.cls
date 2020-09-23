VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Assertions"
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
'    Module: Assertions
'

''
' Provides a set of assertions used to enforce specific comparisons.
'
' @remarks Assertions are used to test for valid values in the user tests.
' <p>To access the assertion methods, use the <b>Assert</b> global function
' which provides all the assertions. The following example shows the
' syntax of using the <b>Assert</b> function.</p>
' <pre>
' Public Sub MyTest()
'     Assert.AreEqual 1, Value, "Value should be 1."
'     ...
' End Sub
' </pre>
' There is no need to create a new object of any type to access the assertions.
'
Option Explicit

Private Const DEF_STRINGCOMPAREMETHOD   As Long = VbCompareMethod.vbBinaryCompare
Private Const DEF_FLOATCOMPARETOLERANCE As Double = 0.00001

Public Enum DateComparePrecision
    ToTheSecond
    ToTheMinute
    ToTheHour
    ToTheDay
    ToTheMonth
    ToTheYear
    ToTheWeek
    ToTheQuarter
    ToTheDayOfYear
    
    #If False Then
        Dim ToTheSecond, ToTheMinute, ToTheHour, ToTheDay, ToTheMonth
        Dim ToTheYear, ToTheWeek, ToTheQuarter, ToTheDayOfYear
    #End If
End Enum


Private mAssertCount            As Long
Private mResult                 As New AssertResult
Private mStringCompareMethod    As VbCompareMethod
Private mFloatCompareTolerance  As Double



''
' Returns the compare method used when comparing strings with <b>Assert.AreEqual</b>.
'
' @return The current method used to compare strings with <b>Assert.AreEqual</b>.
' @remarks This is to allow a specific string comparison method to be applied to
' all <b>Assert.AreEqual</b> calls during testing.
' <p>To use a compare method for a specific assert, use <b>Assert.AreEqualStrings</b>.</p>
' <p>The default compare method is <b>vbCompareBinary</b>.
'
Public Property Get StringCompareMethod() As VbCompareMethod
    StringCompareMethod = mStringCompareMethod
End Property

''
' Sets the compare method used when comparing strings with <b>Assert.AreEqual</b>.
'
' @param RHS The method to be used to compare strings using <b>Assert.AreEqual</b>.
' @remarks This is to allow a specific string comparison method to be applied to
' all <b>Assert.AreEqual</b> calls during testing.
' <p>To use a compare method for a specific assert, use <b>Assert.AreEqualStrings</b>.</p>
' <p>The default compare method is <b>vbCompareBinary</b>.
'
Public Property Let StringCompareMethod(ByVal RHS As VbCompareMethod)
    mStringCompareMethod = RHS
End Property

''
' Returns the tolerance allowed when comparing floats (Double, Single) using <b>Assert.AreEqual</b>.
'
' @return The current tolerance used when comparing floats (Double, Single) using <b>Assert.AreEqual</b>.
' @remarks This is to allow a specific Float comparison method to be applied to
' all <b>Assert.AreEqual</b> calls during testing.
' <p>To use a specific tolerance for a specific assert, use <b>Assert.AreEqualFloats</b>.</p>
' <p>The default tolerance is 0.00001.</p>
'
Public Property Get FloatCompareTolerance() As Double
    FloatCompareTolerance = mFloatCompareTolerance
End Property

''
' Sets the tolerance to be used when comparing Floats (Double, Single) using <b>Assert.AreEqual</b>.
'
' @param RHS The new tolerance to be used when comparing Floats (Double, Single) using <b>Assert.AreEqual</b>.
' @remarks This is to allow a specific Float comparison method to be applied to
' all <b>Assert.AreEqual</b> calls during testing.
' <p>To use a specific tolerance for a specific assert, use <b>Assert.AreEqualFloats</b>.</p>
' <p>The default tolerance is 0.00001.</p>
'
Public Property Let FloatCompareTolerance(ByVal RHS As Double)
    mFloatCompareTolerance = RHS
End Property

''
' This resets all errors and Assert results.
'
' @remarks If an error trap is used during a test then this assert should be called
' if no other assertion is being called within the error trap.
' <p>If another assertion is executed in the error trap, then this method does not
' need to be called.
'
Public Sub Clear()
    Set mResult = AssertResult.Pass
    Call Err.Clear
End Sub

''
' Returns the error code used to identify an assertion that fails.
'
' @return An error code used by assertions.
' @remarks The error code is <b>vbObjectError</b> + 3001.
'
Public Property Get AssertCode() As Long
    AssertCode = ERR_ASSERT_FAIL
End Property

''
' Returns the error code used to identify an assertion
' that is IsIgnored and not cause a failure.
'
' @return Returns the code used to identify an ignore request.
' @remarks The ignore code is <b>vbObjectError</b> + 3002.
'
Public Property Get IgnoreCode() As Long
    IgnoreCode = ERR_ASSERT_IGNORE
End Property

''
' Causes the current test to stop without failing.
'
' @param Message The reason for ignoring the current test.
' @remarks A test that is IsIgnored is still counted as being
' run, however, it is not counted as passing or failing.
'
Public Sub Ignore(Optional ByVal Message As String)
    Call DoAssert(AssertResult.Fail("Ignore", "Assert.Ignore Called", Message, ERR_ASSERT_IGNORE))
End Sub

''
' Causes an immediate assertion failure and stops the current test.
'
' @param Message An user specified message that describes the failure.
' @remarks This assertion forces an immediate failure and stops the test.
' <p>The <b>Err.Number</b> will return the same value as <b>Assert.ErrorCode</b>.</p>
'
Public Sub Fail(Optional ByVal Message As String)
    Call DoAssert(AssertResult.Fail("Fail", "Assert.Fail Called", Message))
End Sub

''
' Asserts that the condition is to be True.
'
' @param Condition The condition to be tested for True.
' @param Message A message that describes the failure.
' @remarks The condition must result in a boolean value. Visual Basic
' can convert most evaluations to a boolean automatically.
'
Public Sub IsTrue(ByVal Condition As Boolean, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsTrue(Condition, Message))
End Sub

''
' Asserts that the condition is False.
'
' @param Condition The condition to be tested for False.
' @param Message A message that describes the failure.
' @remarks Testing for a False condition can be counter-intuitive, therefore,
' care should be taken to test for the correct condition.
' <p>If the condition is True, the assertion will Fail. If the condition
' is False, the condition will pass.</p>
' <pre>
' 'All of these will fail
' Assert.IsFalse True
' Assert.IsFalse 1 = 1
' Assert.IsFalse "hi" = "hi"
'
' 'All of these will pass.
' Assert.IsFalse False
' Assert.IsFalse 1 = 2
' Assert.IsFalse "hi" = "bye"
' </pre>
'
Public Sub IsFalse(ByVal Condition As Boolean, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsFalse(Condition, Message))
End Sub

''
' Asserts that the value is equal to Empty.
'
' @param Value The value to be tested for being Empty.
' @param Message A message that describes the failure.
' @remarks A <b>Message</b> should describe why the assertion failed.
' The framework can provide a reason the assertion failed at the comparison
' level, however, it is usually best to provide a message describing the
' context in which the assertion was testing.
' <p>This uses the <b>VBA.IsEmpty</b> function to test for Empty.</p>
'
Public Sub IsEmpty(ByRef Value As Variant, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsEmpty(Value, Message))
End Sub

''
' Asserts that the value is not Empty.
'
' @param Value The value to be test for not being Empty.
' @param Message A message that describes the failure.
'
'
Public Sub IsNotEmpty(ByRef Value As Variant, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNotEmpty(Value, Message))
End Sub

''
' Asserts that the value is Null.
'
' @param Value The value to be tested for being Null.
' @param Message A user specified message if the assertion fails.
' @remarks Testing for a Null value is not the same as testing
' for a Nothing reference.
'
Public Sub IsNull(ByRef Value As Variant, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNull(Value, Message))
End Sub

''
' Asserts that the value is not Null.
'
' @param Value The value to be tested for not being Null.
' @param Message A user specified message if the assertion fails.
' @remarks Testing for a Null value is not the same as testing
' for a Nothing reference.
'
Public Sub IsNotNull(ByRef Value As Variant, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNotNull(Value, Message))
End Sub

''
' Asserts that the value is Nothing.
'
' @param Value The value to be tested for being Nothing.
' @param Message A user specified message if the assertion fails.
'
Public Sub IsNothing(ByVal Value As IUnknown, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNothing(Value, Message))
End Sub

''
' Asserts that the value is not Nothing.
'
' @param Value The value to be tested for not being Nothing.
' @param Message A user specified message if the assertion fails.
'
Public Sub IsNotNothing(ByVal Value As IUnknown, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNotNothing(Value, Message))
End Sub

''
' Asserts that the Actual value is equal to the Expected value.
'
' @param Expected The value to test for.
' @param Actual The value to test.
' @param Message A message describing why the assertion failed.
' @remarks This is a general case comparison. All numeric types are
' treated as equals when comparing. Doubles can be compared to Longs
' for example.
' <p>String comparisons are dependant on the <b>Assert.StringCompareMethod</b>
' setting. To avoid the default compare method of <b>AreEqual</b> use <b>AreEqualStrings</b>.</p>
' <p>Floats (Doubles, Singles) are compared within the tolerance
' specified by the <b>Assert.FloatCompareTolerance</b> value.</p>
'
Public Sub AreEqual(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As String)
    Call DoAssert(AreEqualAsserter.Assert(Expected, Actual, Message))
End Sub

''
' Asserts that the Actual value is not equal to the NotExpected value.
'
' @param NotExpected The value not to have.
' @param Actual The value to test.
' @param Message A message describing why the assertion failed.
' @remarks This is a general case comparison. All numeric types are
' treated as equals when comparing. Doubles can be compared to Longs
' for example.
' <p>String comparisons are dependant on the <b>Assert.StringCompareMethod</b> setting.</p>
' <p>Floats (Doubles, Singles) are compared within the tolerance
' specified by the <b>Assert.FloatCompareTolerance</b> value.</p>
'
Public Sub AreNotEqual(ByRef NotExpected As Variant, ByRef Actual As Variant, Optional ByVal Message As String)
    Call DoAssert(AreNotEqualAsserter.Assert(NotExpected, Actual, Message))
End Sub

''
' Asserts that two object variables point to the same object.
'
' @param Expected The object to test for.
' @param Actual The object being tested.
' @param Message A message describing why the assertion failed.
'
Public Sub AreSame(ByVal Expected As IUnknown, ByVal Actual As IUnknown, Optional ByVal Message As String)
    Call DoAssert(EqualityAsserter.AreSame(Expected, Actual, Message))
End Sub

''
' Asserts that two object variables do not point to the same object.
'
' @param NotExpected The object not to have.
' @param Actual The object to test.
' @param Message A message describing why the assertion failed.
'
Public Sub AreNotSame(ByVal NotExpected As IUnknown, ByVal Actual As IUnknown, Optional ByVal Message As String)
    Call DoAssert(EqualityAsserter.AreNotSame(NotExpected, Actual, Message))
End Sub

''
' Asserts that two dates are equal within a specific precision.
'
' @param Expected The expected date value.
' @param Actual The date to be tested.
' @param Precision The precision to compare the two dates to for equality.
' @param Message A message describing why the assertion failed.
' @remarks When dates are compared, they must be uqual to the specified precision, not
' just within 1 unit of the specified precision. For example, if the precision is
' to the minute, then both values must be the same down to the minute. The two values
' cannot just be within 1 minute of eachother to be equal.
'
Public Sub AreEqualDates(ByRef Expected As Date, ByRef Actual As Date, ByVal Precision As DateComparePrecision, Optional ByVal Message As String)
    Call DoAssert(AreEqualDatesAsserter.Assert(Expected, Actual, Precision, Message))
End Sub

''
' Asserts that two floats are equal within the same tolerance.
'
' @param Expected The expected value.
' @param Actual The value to be tested.
' @param Tolerance The maximum value the two values can be different and remain equal.
' @param Message A message describing why the assertion failed.
'
Public Sub AreEqualFloats(ByRef Expected As Double, ByRef Actual As Double, Optional ByVal Tolerance As Double = 0.00001, Optional ByVal Message As String)
    Call DoAssert(EqualityAsserter.AreEqualFloats(Expected, Actual, Tolerance, Message))
End Sub

''
' Asserts that two strings are equal using the specific compare method.
'
' @param Expected The expected value.
' @param Actual The value to be tested.
' @param CompareMethod How to compare the two strings.
' @param Message A message describing why the asseriton failed.
'
Public Sub AreEqualStrings(ByRef Expected As String, ByRef Actual As String, Optional ByVal CompareMethod As VbCompareMethod = vbTextCompare, Optional ByVal Message As String)
    Call DoAssert(EqualityAsserter.AreEqualStrings(Expected, Actual, CompareMethod, Message))
End Sub

''
' Asserts that a collection or array contains a specific value.
'
' @param Value The value to find in the container.
' @param Container The list to search for the value.
' @param Message A message describing why the assertion failed.
' @remarks The container can be an Array or an object that supports enumeration.
'
Public Sub Contains(ByRef Value As Variant, ByRef Container As Variant, Optional ByVal Message As String)
    Call DoAssert(ContainsAsserter.Assert(Value, Container, Message))
End Sub

''
' Asserts that LHS is less than RHS.
'
' @param LHS The Left-Hand-Side of the comparison.
' @param RHS The Right-Hand-Side of the comparison.
' @param Message A message describing why the assertion failed.
'
Public Sub Less(ByRef LHS As Variant, ByRef RHS As Variant, Optional ByVal Message As String)
    Call DoAssert(ComparisonAsserter.Less(LHS, RHS, Message))
End Sub

''
' Asserts that LHS is greater than RHS.
' @param LHS The Left-Hand-Side of the comparison.
' @param RHS The Right-Hand-Side of the comparison.
' @param Message A message describing why the assertion failed.
'
Public Sub Greater(ByRef LHS As Variant, ByRef RHS As Variant, Optional ByVal Message As String)
    Call DoAssert(ComparisonAsserter.Greater(LHS, RHS, Message))
End Sub

''
' Performs a simple expression evaluation using the VB <b>Like</b> operator.
'
' @param Expression The <b>Like</b> expression.
' @param Value The value to be evaluated using Like.
' @param Message A message describing why the assertion failed.
'
Public Sub IsLike(ByVal Expression As String, ByVal Value As String, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsLike(Expression, Value, Message))
End Sub

''
' Performs a simple expression evaluation using the VB <b>Like</b> operator.
'
' @param Expression The <b>Like</b> expression.
' @param Value The value to be evaluated using Like.
' @param Message A message describing why the assertion failed.
'
Public Sub IsNotLike(ByVal Expression As String, ByVal Value As String, Optional ByVal Message As String)
    Call DoAssert(ConditionAsserter.IsNotLike(Expression, Value, Message))
End Sub

''
' Performs a simple expression evaluation using the VB <b>Like</b> operator.
'
' @param Expression The <b>Like</b> expression.
' @param Value The value to be evaluated using Like.
' @param Message A message describing why the assertion failed.
' @remarks The <b>Like</b> uses <b>Option Compare Text</b> and is
' therefore not case sensitive.
'
Public Sub IsLikeNoCase(ByVal Expression As String, ByVal Value As String, Optional ByVal Message As String)
    Call DoAssert(IsLikeNoCaseAsserter.IsLikeNoCase(Expression, Value, Message))
End Sub

''
' Performs a simple expression evaluation using the VB <b>Like</b> operator.
'
' @param Expression The <b>Like</b> expression.
' @param Value The value to be evaluated using Like.
' @param Message A message describing why the assertion failed.
' @remarks The <b>Like</b> uses <b>Option Compare Text</b> and is
' therefore not case sensitive.
'
Public Sub IsNotLikeNoCase(ByVal Expression As String, ByVal Value As String, Optional ByVal Message As String)
    Call DoAssert(IsLikeNoCaseAsserter.IsNotLikeNoCase(Expression, Value, Message))
End Sub

''
' Searches for the value in the parameter list.
'
' @param Value the Value to be searched for.
' @param List The list to be searched.
'
Public Sub IsInList(ByRef Value As Variant, ParamArray List() As Variant)
    Dim v() As Variant
    v = List
    Call DoAssert(IsInListAsserter.IsInList(Value, v))
End Sub

''
' Verifies that the value is not in the list.
'
' @param Value The value that should not be in the list.
' @param List The list to verify the value is not in.
'
Public Sub IsNotInList(ByRef Value As Variant, ParamArray List() As Variant)
    Dim v() As Variant
    v = List
    Call DoAssert(IsInListAsserter.IsNotInList(Value, v))
End Sub

''
' Sends a message to all listeners.
'
' @param Message The message to be sent.
' @param Args A list of arguments that can be sent to listeners to use.
' @remarks This does not cause a test to fail.
'
Public Sub Say(ByVal Message As String, ParamArray Args() As Variant)
    Dim v() As Variant
    v = Args
    Call Listener.Hear(Message, v)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' This allows test methods to get results of failed assertions.
'
' @return The result of the failed assertion.
' @remarks The result can only be retrieved once. After the
' result has been retrieved, it is cleared internally and will
' return Nothing until an assertion fails again.
'
Friend Property Get Result() As AssertResult
    Set Result = mResult
End Property

Friend Property Set Result(ByVal RHS As AssertResult)
    Set mResult = RHS
End Property

Friend Function SetAssertCount(ByVal Count As Long) As Long
    SetAssertCount = mAssertCount
    mAssertCount = Count
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncrementAssertCount()
    mAssertCount = mAssertCount + 1
End Sub

Private Sub DoAssert(ByVal Result As AssertResult)
    Call Clear
    Call IncrementAssertCount
    Call Result.NotifyOnFailure
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mStringCompareMethod = DEF_STRINGCOMPAREMETHOD
    mFloatCompareTolerance = DEF_FLOATCOMPARETOLERANCE
End Sub

