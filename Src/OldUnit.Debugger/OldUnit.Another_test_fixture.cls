VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Another_test_fixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim testNames As Collection
Dim helper As TestFixtureHelper

'**********************************
' Must be implemented to work with
' OldUnit testrunner
' Should return collection of function names
' to be executed by the test runner
'**********************************
Public Property Get Tests() As Collection
  Set Tests = testNames
End Property

'**********************************
' Must be implemented to work with
' OldUnit testrunner
' Should return a new instance of
' the testfixture
'**********************************
Public Function Create() As TestFixture
  Set Create = New TestFixture
End Function

'**********************************
' The names of all test functions
' must be manually added to the
' testNames collection when class
' is initialized
'
' Also it is adivsable to add a
' TestFixtureHelper to simplyfy
' the test functions
'*********************************
Private Sub Class_initialize()
  Set testNames = New Collection
  Set helper = New TestFixtureHelper
  Call testNames.Add("This_test_should_fail")
  Call testNames.Add("This_test_should_pass")
End Sub

'*********************************
' A test Method is a function that
' returns Nothing if no error is raised
' or returns a testErr object should
' an error have been raised
'*********************************
Public Function This_test_should_fail() As TestErr
On Error GoTo errorHandler
  Call Err.Raise(vbObjectError + 1000, "This test should fail", "Something went horribly wrong")
  Set This_test_should_fail = Nothing
Exit Function

errorHandler:
  Set This_test_should_fail = helper.fail
End Function

'*********************************
' A test Method is a function that
' returns Nothing if no error is raised
' or returns a testErr object should
' an error have been raised
'*********************************
Public Function This_test_should_pass() As TestErr
  On Error GoTo errorHandler
  Set This_test_should_pass = Nothing
Exit Function

errorHandler:
  This_test_should_pass = helper.fail
End Function

Private Sub Class_terminate()
  Set testNames = Nothing
  Set helper = Nothing
End Sub

