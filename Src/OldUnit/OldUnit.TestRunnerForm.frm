VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TestRunnerForm 
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16548
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   16548
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRunSelected 
      Caption         =   "Run selected"
      Height          =   312
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdRunAll 
      Caption         =   "Run all"
      Height          =   312
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox tbTestDetails 
      Height          =   2652
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6840
      Width           =   16332
   End
   Begin MSComctlLib.TreeView trvTestList 
      Height          =   6132
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   16332
      _ExtentX        =   28808
      _ExtentY        =   10816
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "TestRunnerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim localTestRunner As TestRunner

Private Const ISFIXTURE As Integer = -1
Private Const KEYSPLITSTRING As String = "<--SPLIT-->"
Private Const METHODNAME As String = "methodname"
Private Const FIXTURENAME As String = "fixturename"

Private Sub cmdRunAll_Click()
  Call localTestRunner.RunAll
End Sub

Public Sub NewResult(result As TestResult)
  Dim resultNode As Node
  Set resultNode = trvTestList.Nodes(CreateTestMethodKey(result.FIXTURENAME, result.Name))
  resultNode.Text = resultNode.Text + " (" + IIf(result.HasPassed, "Passed", "Failed") + ")"
  resultNode.Tag = result.failureText
  
End Sub

Private Sub cmdRunSelected_Click()
  RunSelected
End Sub

Private Sub Form_Load()
  trvTestList.Nodes.Clear
End Sub

Public Sub LoadTests(newTestRunner As TestRunner)
  Dim testFixture As Variant
  Dim testMethod As Variant
  Dim newNode As Node
  Dim fixtureKey As String
  Dim testMethodKey As String
  Dim testMethodText As String
  
  Set localTestRunner = newTestRunner
  Call localTestRunner.AddReporter(Me)
  
  For Each testFixture In localTestRunner
    fixtureKey = TypeName(testFixture)
    Set newNode = trvTestList.Nodes.Add(, , fixtureKey, fixtureKey)
    newNode.Tag = ISFIXTURE
    newNode.Expanded = True
    For Each testMethod In testFixture.Tests
      testMethodKey = CreateTestMethodKey(fixtureKey, CStr(testMethod))
      testMethodText = CreateTestMethodText(CStr(testMethod))
      Set newNode = trvTestList.Nodes.Add(fixtureKey, tvwChild, testMethodKey, testMethodText)
    Next
  Next
End Sub

Private Function CreateTestMethodKey(testFixtureName As String, testMethodName As String)
  CreateTestMethodKey = testFixtureName + KEYSPLITSTRING + testMethodName
End Function

Private Function CreateTestMethodText(testMethodName As String)
  CreateTestMethodText = Replace(testMethodName, "_", " ")
End Function

Private Sub trvTestList_NodeCheck(ByVal Node As MSComctlLib.Node)
  Dim childNode As Node
  
End Sub

Private Sub trvTestList_NodeClick(ByVal Node As MSComctlLib.Node)
  tbTestDetails.Text = Node.Tag
End Sub

Private Function DecomposeTestMethodKey(testMethodKey As String) As Collection
  Dim result As Collection
  Dim splitArray() As String
  Set result = New Collection
  
  splitArray = Split(testMethodKey, KEYSPLITSTRING)
  Call result.Add(splitArray(0), FIXTURENAME)
  Call result.Add(splitArray(1), METHODNAME)
  Set DecomposeTestMethodKey = result
End Function

Private Sub RunSelected()
  Dim testNode As Node
  Dim decomposedKey As Collection
  
  For Each testNode In trvTestList.Nodes
    If testNode.Checked And Not testNode.Tag = ISFIXTURE Then
      Set decomposedKey = DecomposeTestMethodKey(testNode.Key)
      Call localTestRunner.RunTest(decomposedKey(FIXTURENAME), decomposedKey(METHODNAME))
    End If
  Next
End Sub
