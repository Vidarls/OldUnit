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
Dim results                 As Collection

Public Sub AddFixture(testFixture As Variant)
    Call testFixtures.Add(testFixture)
End Sub

Public Sub AddReporter(reporter As Variant)

End Sub

Public Sub Run()
    'On Error GoTo handleRunnerErrors
    Dim testFixture As Variant
    Dim instanceUnderTest As Variant
    Dim testName As Variant
    Dim testFailed As TestErr
    
    For Each testFixture In testFixtures
        For Each testName In testFixture.Tests
            Set instanceUnderTest = testFixture.Create()
            Set testFailed = CallByName(instanceUnderTest, testName, VbMethod)
            If Not testFailed Is Nothing Then
                Call ReportFailure(TypeName(testFixture), CStr(testName), testFailed.AsString)
            Else
                Call ReportPass(TypeName(testFixture), CStr(testName))
            End If
        Next
    Next
    
    Call CreateReports
Exit Sub

handleRunnerErrors:
  MsgBox "Testrunner internal error: " + CStr(Err.Number) + " " + Err.Description
End Sub

Private Function FindOrCreateTestFixtureResults(testFixtureName As String)
    Dim testFixtureResult As TestFixtureResults

    For Each testFixtureResult In results
        If testFixtureResult.fixtureName = testFixtureName Then
            Set FindOrCreateTestFixtureResults = testFixtureResult
            Exit Function
        End If
    Next
    
    Set FindOrCreateTestFixtureResults = New TestFixtureResults
End Function

Private Sub CreateReports()
    Dim reporter As Variant
    For Each reporter In reporters
        Call reporter.CreateReport(results)
    Next
End Sub

Private Sub ReportFailure(fixtureName As String, testName As String, failureText As String)
    Dim testFixtureResult As TestFixtureResults
    Dim newTestResult As TestResult
    
    Set testFixtureResult = FindOrCreateTestFixtureResults(fixtureName)
    Set newTestResult = New TestResult
    Call newTestResult.Failed(fixtureName, testName, failureText)
    Call testFixtureResult.AddTestResult(newTestResult)
    Call UpdateTestResults(testFixtureResult)
End Sub

Private Sub UpdateTestResults(testFixtureResult As TestFixtureResults)
    Dim existingTestFixtureResult As TestFixtureResults
    Dim newTestResults As Collection: Set newTestResults = New Collection
    Dim foundAndReplaced As Boolean: foundAndReplaced = False
    For Each existingTestFixtureResult In results
        If existingTestFixtureResult.fixtureName = testFixtureResult.fixtureName Then
            Call newTestResults.Add(testFixtureResult)
            foundAndReplaced = True
        Else
            Call newTestResults.Add(existingTestFixtureResult)
        End If
    Next
    If Not foundAndReplaced Then Call newTestResults.Add(testFixtureResult)
    
    Set results = newTestResults
End Sub


Private Sub ReportPass(fixtureName As String, testName As String)
    Dim testFixtureResult As TestFixtureResults
    Dim newTestResult As TestResult
    
    Set testFixtureResult = FindOrCreateTestFixtureResults(fixtureName)
    Set newTestResult = New TestResult
    Call newTestResult.Passed(fixtureName, testName)
    Call testFixtureResult.AddTestResult(newTestResult)
    Call UpdateTestResults(testFixtureResult)
End Sub

Private Sub Class_initialize()
    Set testFixtures = New Collection
    Set results = New Collection
    Set reporters = New Collection
    Call reporters.Add(New GuiReporter)
End Sub

Private Sub Class_terminate()
    Set testFixtures = Nothing
    Set reporters = Nothing
    Set results = Nothing
End Sub

