** DebugWindowListener **

This is a simple ITestListener that can be used to output results in the Immediate debug window.

To use this class, add it to your project that will run the tests. When the Run method is is to be called on the test, pass in an instance of DebugWindowListener.

Example on the usage:

   Dim Tests As New TestSuite
   Tests.Add New TestClass1
   ... Add all the test classes.

   Tests.Run New DebugWindowListener


This will output the results to the Immediate window.

