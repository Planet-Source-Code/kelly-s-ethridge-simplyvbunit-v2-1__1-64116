Attribute VB_Name = "modStaticClasses"
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
'    Module: modStaticClasses
'

Option Explicit

Public Sim              As New Constructors
Public Assert           As New Assertions
Public NullTestListener As New NullTestListenerStatic
Public TestContext      As New TestContextStatic
Public AssertResult     As New AssertResultStatic
Public Comparer         As New ComparerStatic
Public TestComparer     As New TestComparerStatic
Public Iterations       As New IterationStack

' Assertion helpers
Public ConditionAsserter        As New ConditionAsserter
Public AreEqualAsserter         As New AreEqualAsserter
Public AreNotEqualAsserter      As New AreNotEqualAsserter
Public EqualityAsserter         As New EqualityAsserter
Public AreEqualDatesAsserter    As New AreEqualDatesAsserter
Public ContainsAsserter         As New ContainsAsserter
Public ComparisonAsserter       As New ComparisonAsserter
Public IsLikeNoCaseAsserter     As New IsLikeNoCaseAsserter
Public EqualAsserter            As New EqualAsserter
Public IsInListAsserter         As New IsInListAsserter
