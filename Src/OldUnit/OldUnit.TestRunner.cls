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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim testFixtures            As Collection
Dim reporters               As Collection

Public Sub AddFixture(testFixture As Variant)
  Call testFixtures.Add(testFixture, TypeName(testFixture))
End Sub

Public Sub AddReporter(reporter As Variant)
  Call reporters.Add(reporter)
End Sub

Public Sub Show(Optional modal As Integer = 0)
  Load TestRunnerForm
  Call TestRunnerForm.LoadTests(Me)
  Call TestRunnerForm.Show(modal)
End Sub

Public Sub ShowAndRun(Optional modal As Integer = 0)
  Load TestRunnerForm
  Call TestRunnerForm.LoadTests(Me)
  RunAll
  Call TestRunnerForm.Show(modal)
End Sub

Public Sub RunAll()
'On Error GoTo handleRunnerErrors
  Dim testFixture As Variant
  Dim instanceUnderTest As Variant
  Dim testName As Variant
  
  For Each testFixture In testFixtures
    For Each testName In testFixture.Tests
      Call RunTest(TypeName(testFixture), CStr(testName))
    Next
  Next
Exit Sub

handleRunnerErrors:
  MsgBox "Testrunner internal error: " + CStr(Err.Number) + " " + Err.Description
End Sub

Public Sub RunTest(testFixtureName As String, testMethodName As String)
'On Error GoTo handleRunnerErrors
  Dim instanceUnderTest As Variant
  Dim testFailed As TestErr
  
  Set instanceUnderTest = testFixtures(testFixtureName).Create
  Set testFailed = CallByName(instanceUnderTest, testMethodName, VbMethod)
  
  If testFailed Is Nothing Then
    Call ReportPass(testFixtureName, testMethodName)
  Else
    Call ReportFailure(testFixtureName, testMethodName, testFailed)
  End If
Exit Sub

handleRunnerErrors:
  MsgBox "Testrunner internal error: " + CStr(Err.Number) + " " + Err.Description
End Sub

Private Sub ReportFailure(testFixtureName As String, testName As String, testFailed As TestErr)
  Dim newTestResult As TestResult
  Dim failureText As String
  
  failureText = "Test failed" + vbNewLine + _
                CStr(testFailed.Number) + vbNewLine + _
                testFailed.Source + vbNewLine + vbNewLine + _
                testFailed.Description
                
  Set newTestResult = New TestResult
  Call newTestResult.Failed(testFixtureName, testName, failureText)
  Call ReportResultToReporters(newTestResult)
End Sub

Private Sub ReportPass(testFixtureName As String, testName As String)
  Dim newTestResult As TestResult

  Set newTestResult = New TestResult
  Call newTestResult.Passed(testFixtureName, testName)
  Call ReportResultToReporters(newTestResult)
End Sub

Private Sub ReportResultToReporters(result As TestResult)
  Dim reporter As Variant
  For Each reporter In reporters
    Call reporter.NewResult(result)
  Next
End Sub

Private Sub Class_initialize()
  Set testFixtures = New Collection
  Set reporters = New Collection
End Sub

Private Sub Class_terminate()
  Set testFixtures = Nothing
  Set reporters = Nothing
End Sub

'TestRunner as custom collection implementation
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = testFixtures.[_NewEnum]
End Function
