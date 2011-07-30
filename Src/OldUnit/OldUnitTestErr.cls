VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim internalNumber      As Long
Dim internalSource      As String
Dim internalDescription As String

Public Property Get Number() As Long
  Number = internalNumber
End Property

Public Property Get Source() As String
  Source = internalSource
End Property

Public Property Get Description() As String
  Description = internalDescription
End Property

Public Property Get AsString() As String
  AsString = CStr(internalNumber) + " : " + internalSource + " : " + internalDescription
End Property


Public Sub CopyFromErr()
  internalNumber = Err.Number
  internalSource = Err.Source
  internalDescription = Err.Description
End Sub

