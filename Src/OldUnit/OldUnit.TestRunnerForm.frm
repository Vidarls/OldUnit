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
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   312
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "TestRunnerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim localTestRunner As TestRunner

Private Sub Form_Load()
  trvTestList.Nodes.Clear
End Sub

Public Sub LoadTests(newTestRunner As TestRunner)
  Dim test As Variant
  Dim newNode As Node
  Dim testFixtureIndex As Integer
  Dim treeViewIndex As Integer
  
  Set localTestRunner = newTestRunner
  
  For Each test In localTestRunner
    
    
  Next
End Sub
