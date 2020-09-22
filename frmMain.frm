VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bl@ck$corpion"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Contains"
      Height          =   3555
      Left            =   60
      TabIndex        =   9
      Top             =   2280
      Width           =   7695
      Begin VB.CommandButton cmdBrowsePicture 
         Caption         =   "..."
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label lblX 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture:"
         Height          =   255
         Index           =   3
         Left            =   3540
         TabIndex        =   12
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblX 
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   375
      End
      Begin VB.Image imgPic 
         Height          =   2880
         Left            =   3480
         Picture         =   "frmMain.frx":0442
         Stretch         =   -1  'True
         Top             =   540
         Width           =   4110
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Password"
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   7695
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   8
         Text            =   "password"
         Top             =   300
         Width           =   2295
      End
      Begin VB.CheckBox chkPassword 
         Caption         =   "Password protected"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   1455
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtHeader 
         Height          =   285
         Left            =   1140
         TabIndex        =   18
         Text            =   "Hello"
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1140
         TabIndex        =   4
         Text            =   "DummyApp"
         Top             =   660
         Width           =   6435
      End
      Begin VB.TextBox txtExeFile 
         Height          =   285
         Left            =   1140
         TabIndex        =   2
         Text            =   "DummyApp.exe"
         Top             =   240
         Width           =   6435
      End
      Begin VB.Label lblX 
         BackStyle       =   0  'Transparent
         Caption         =   "Header:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label lblX 
         Caption         =   "Exe caption :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblX 
         Caption         =   "Exe name :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile!!!!"
      Default         =   -1  'True
      Height          =   435
      Left            =   6000
      TabIndex        =   0
      Top             =   6000
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2220
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE VOTE!!!!"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   15
      Top             =   6360
      Width           =   435
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "created by Wouter van Koppen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   3060
      TabIndex        =   14
      Top             =   6360
      Width           =   2685
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   645
      TabIndex        =   13
      Top             =   6360
      Width           =   2205
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Read readme.txt for info
'Title:      Bl@ck$corpion
'Category:   Exe compiler
'Author:     Wouter van Koppen
'Mail:       Xbrain3000@ hotmail.com

Private Sub cmdBrowsepicture_Click()
    On Error GoTo Error
    Dialog.CancelError = True
    Dialog.ShowOpen
    imgPic.Picture = LoadPicture(Dialog.FileName)
Error:
End Sub

Private Sub cmdCompile_Click()
'This copies first the interface "BlackScorpion.dat"
'to your new exe.
'Then it writes the picture and text what you choosed.

    On Local Error GoTo Error
    Dim PropertyBuilder As New PropertyBag
    Dim Temp As Variant
    Dim StartPosition As Long

    With PropertyBuilder
        .WriteProperty "Caption", txtCaption.Text
        .WriteProperty "Text", txtText.Text
        .WriteProperty "Picture", imgPic.Picture
        .WriteProperty "Protected", chkPassword.Value
        .WriteProperty "Password", txtPassword.Text
        .WriteProperty "Header", txtHeader.Text
        'Here the data will be writed in the exe file
    End With
    
   'Create the main interface stored in BlackScorpion.dat
    FileCopy App.Path & "\BlackScorpion.dat", App.Path & "\" & txtExeFile.Text
    
    'Add the picture and text
    Open App.Path & "\" & txtExeFile.Text For Binary As #1
        StartPosition = LOF(1)   'When traced we can add the data
        Temp = PropertyBuilder.Contents
        Seek #1, LOF(1)
        Put #1, , Temp   'Write data
        Put #1, , StartPosition  'Write starting point
    Close #1
    
    'No errors!
    MsgBox "Executable created! ", vbInformation
    Exit Sub
    
    'Done
    
Error:
    If Err.Number = 53 Then
        'Main interface not found
        MsgBox "BlackScorpion.dat is not found!", vbExclamation
        Exit Sub
    End If
    'Other error
    Msg = "There was an error during compilation" & vbCrLf
    Msg = Msg & vbCrLf & Err.Description
    MsgBox Msg, vbCritical, "Error"
End Sub

Private Sub Form_Load()
    MsgBox "Apply <app.path & ""\Exe interface\Exe interface.vbp""> to learn how this works.", vbInformation
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = &H0
End Sub

Private Sub lblmail_Click()
    Shell "start mailto:Xbrain3000@hotmail.com"
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = &HFF0000
End Sub

Private Sub chkPassword_Click()
    txtPassword.Enabled = chkPassword.Value
End Sub

