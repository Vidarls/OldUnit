VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Assert_equal_tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim localTests As Collection
Dim helper As TestFixtureHelper
Dim Assert As Asserter

Const ARBITRARY_ERROR_CODE As Long = vbObjectError

Public Property Get Tests() As Collection
  Set Tests = localTests
End Property

Public Function Create() As Assert_equal_tests
  Set Create = New Assert_equal_tests
End Function

Private Function Fail() As TestErr
  Set Fail = helper.Fail
End Function

Private Sub Class_initialize()
  Set localTests = New Collection
  Set Assert = New Asserter
  Set helper = New TestFixtureHelper
  
  localTests.Add "One_should_equal_one"
  localTests.Add "One_should_not_equal_two"
  localTests.Add "Text_should_equal_Text"
  localTests.Add "Text_should_not_equal_text"
  localTests.Add "True_should_equal_true"
  localTests.Add "True_should_not_equal_false"
End Sub


Public Function One_should_equal_one()
On Error GoTo failed
  Call Assert.That(1).Equals(1)
  Set One_should_equal_one = Nothing
Exit Function

failed:
  Set One_should_equal_one = Fail
End Function

Public Function One_should_not_equal_two() As TestErr
On Error GoTo reportError
  Assert.That(1).Equals (2)
  Err.Raise ARBITRARY_ERROR_CODE, "One_should_not_equal_two", "Assertion did not fail"
Exit Function

reportError:
  If Err.Number = Assert.AssertionErrorNumber Then
    Set One_should_not_equal_two = Nothing
    Exit Function
  End If
  Set One_should_not_equal_two = Fail
End Function

Public Function Text_should_equal_Text()
On Error GoTo failed
  Assert.That("Text").Equals ("Text")
  Set Text_should_equal_Text = Nothing
Exit Function

failed:
  Set Text_should_equal_Text = Fail
End Function

Public Function Text_should_not_equal_text()
On Error GoTo failed
  Assert.That("Text").Equals ("text")
  Err.Raise ARBITRARY_ERROR_CODE, "Text_should_not_equal_text", "Assertion did not fail"
Exit Function

failed:
  If Err.Number = Assert.AssertionErrorNumber Then
    Set Text_should_not_equal_text = Nothing
    Exit Function
  End If
  Set Text_should_not_equal_text = Fail
End Function

Public Function True_should_equal_true()
On Error GoTo failed
  Assert.That(True).Equals (True)
  Set True_should_equal_true = Nothing
Exit Function

failed:
  Set True_should_equal_true = Fail
End Function

Public Function True_should_not_equal_false()
On Error GoTo failed
  Assert.That(True).Equals (False)
  Err.Raise ARBITRARY_ERROR_CODE, "True_should_not_equal_false", "Assertion did not fail"
Exit Function

failed:
  If Err.Number = Assert.AssertionErrorNumber Then
    Set True_should_not_equal_false = Nothing
    Exit Function
  End If
  Set True_should_not_equal_false = Fail
End Function

Private Sub Class_terminate()
  Set localTests = Nothing
  Set Assert = Nothing
End Sub


