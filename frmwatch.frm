VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Stop Watch"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4320
      Top             =   4320
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H008080FF&
      Caption         =   "Stop"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFFF80&
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdPaush 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pause"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000FF00&
      Caption         =   "Start"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "M.Sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Hour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPause_Click()
    Timer1.Enabled = False
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdPaush_Click()
  Timer1.Enabled = False
End Sub

Private Sub cmdReset_Click()
    Timer1.Enabled = False
    Label1.Caption = "00"
    Label2.Caption = "00"
    Label3.Caption = "00"
    Label4.Caption = "00"
End Sub

Private Sub cmdStart_Click()
    Timer1.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
End Sub



Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
    Label4.Caption = Val(Label4.Caption) + 1
    If Label4.Caption = 60 Then
         Label4.Caption = 0
         Label3.Caption = Val(Label3.Caption) + 1
         If Label3.Caption = 60 Then
            Label3.Caption = 0
            Label2.Caption = Val(Label2.Caption) + 1
             If Label2.Caption = 60 Then
                Label2.Caption = 0
                Label2.Caption = Val(Label2.Caption) + 1
                If Label2.Caption = 60 Then
                  Label2.Caption = 0
                End If
            End If
        End If
      End If
End Sub

