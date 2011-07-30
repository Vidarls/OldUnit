VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8508
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8508
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Run tests"
      Height          =   3252
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************
'Setupcode for testrunner
'aswell as the run command
'************************
Private Sub Command1_Click()
    Dim runner As TestRunner
    Set runner = New TestRunner
    Call runner.AddFixture(New TestFixture)
    Call runner.Run
    
    
End Sub

Private Sub Form_Load()

End Sub
