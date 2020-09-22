VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   2475
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   2055
      Begin VB.Label lblInfo 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1980
         Width           =   1515
      End
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled Executable"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "created with Bl@ck$corpion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   540
      Width           =   3135
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wouter van Koppen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   1260
      Width           =   3135
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Xbrain3000@hotmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1500
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMail.ForeColor = &H0
End Sub

Private Sub lblmail_Click()
    Shell "start mailto:Xbrain3000@hotmail.com"
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMail.ForeColor = &HFF0000
End Sub

