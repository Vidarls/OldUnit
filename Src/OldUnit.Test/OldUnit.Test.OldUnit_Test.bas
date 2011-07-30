Attribute VB_Name = "OldUnit_Test"
Public Sub Main()
  Dim runner As New TestRunner
  Set runner = New TestRunner
  Call runner.AddFixture(New Assert_equal_tests)
  Call runner.ShowAndRun(1)
End Sub

Public Sub Something()

End Sub
