VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim internalFixtureName     As String
Dim internalName            As String
Dim internalPassed          As Boolean
Dim internalFailureText     As String

Public Property Get fixtureName() As String
  fixtureName = internalFixtureName
End Property

Public Property Get Name() As String
  Name = internalName
End Property

Public Property Get HasPassed() As Boolean
  HasPassed = internalPassed
End Property

Public Property Get failureText() As String
  failureText = internalFailureText
End Property

Public Sub Passed(fixtureName As String, testName As String)
  internalFixtureName = fixtureName
  internalName = testName
  internalPassed = True
  internalFailureText = "Passed"
End Sub

Public Sub Failed(fixtureName As String, testName As String, failureText As String)
  internalFixtureName = fixtureName
  internalName = testName
  internalPassed = False
  internalFailureText = failureText
End Sub
