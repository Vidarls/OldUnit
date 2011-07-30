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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim localItemToAssertOn As Variant
Const ASSERTION_ERROR_CODE As Long = 1000 + vbObjectError

Friend Sub AssignItemToAssertOn(item As Variant)
  If IsObject(item) Then
    Set localItemToAssertOn = item
  Else
   localItemToAssertOn = item
  End If
End Sub

Public Sub Equals(item As Variant)
  If item = localItemToAssertOn Then Exit Sub
  
  Dim errorText As String
  errorText = "Items not equal" + vbNewLine + _
              "Value: " + CStr(localItemToAssertOn) + vbNewLine + _
              "Expected: " + CStr(item)
  Err.Raise ASSERTION_ERROR_CODE, "Assert.Equals", errorText
End Sub
