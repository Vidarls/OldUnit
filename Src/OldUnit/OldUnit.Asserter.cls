VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asserter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ASSERTION_ERROR_CODE As Long = 1000 + vbObjectError

Public Property Get AssertionErrorNumber() As Long
  AssertionErrorNumber = ASSERTION_ERROR_CODE
End Property

Public Function That(item As Variant) As Assertions
On Error GoTo handler
  Dim newAssertions As Assertions
  Set newAssertions = New Assertions
  Call newAssertions.AssignItemToAssertOn(item)
  Set That = newAssertions
Exit Function

handler:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
