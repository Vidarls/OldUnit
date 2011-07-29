VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFixtureResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim internalTestResults         As Collection
Dim internalFixtureName         As String
Dim internalNumberOfFailures    As Integer

Public Property Get TestResults() As Collection
    Set TestResults = internalTestResults
End Property

Public Property Get fixtureName() As String
    fixtureName = internalFixtureName
End Property

Public Property Get HasFailures() As Boolean
    HasFailures = internalNumberOfFailures > 0
End Property

Public Property Get NumberOfFailures() As Integer
    NumberOfFailures = internalNumberOfFailures
End Property

Public Sub AddTestResult(newTestResult As testResult)
    internalFixtureName = newTestResult.fixtureName
    If Not newTestResult.HasPassed Then
        internalNumberOfFailures = internalNumberOfFailures + 1
    End If
    Call internalTestResults.Add(newTestResult)
End Sub

Private Sub Class_initialize()
    Set internalTestResults = New Collection
    internalNumberOfFailures = 0
End Sub

Private Sub Class_terminate()
    Set internalTestResults = Nothing
End Sub
