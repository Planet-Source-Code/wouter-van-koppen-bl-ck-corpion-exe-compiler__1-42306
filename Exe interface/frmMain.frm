VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   180
      Width           =   3915
   End
   Begin VB.Label lblText 
      Height          =   2895
      Left            =   4200
      TabIndex        =   4
      Top             =   180
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
      Left            =   4200
      TabIndex        =   3
      Top             =   4560
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
      Left            =   4200
      TabIndex        =   2
      Top             =   4320
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
      Left            =   4200
      TabIndex        =   1
      Top             =   3600
      Width           =   3135
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
      Left            =   4200
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Image imgPic 
      Height          =   4200
      Left            =   60
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3990
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PropBag As New PropertyBag      'The property bag

'This is the interace of the executable.

'Feel free to steal, modify and redistribute, but
'please give me a VOTE!!!

Private Sub Form_Load()
    On Local Error Resume Next
    Dim BeginPos As Long
    Dim varTemp As Variant
    Dim Infile As Integer
    
    Dim byteArr() As Byte
    
    Infile = FreeFile
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #Infile
        Get #1, LOF(1) - 3, BeginPos    'Get the start position of data

        Seek #1, BeginPos               'Seek to start
        Get #1, , varTemp               'Get the property bag
        
        byteArr = varTemp
        PropBag.Contents = byteArr      'Load the property bag
    
        PropBag.WriteProperty "LOF", LOF(1)
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #Infile
        
    If Val(PropBag.ReadProperty("Protected", "0")) > 0 Then
        Dim PasswordInput As String
        
        PasswordInput = InputBox("Enter password:", "Password required")
        
        If PasswordInput <> PropBag.ReadProperty("Password") Then
            MsgBox "Invalid password, try again.", vbCritical, "Bad shot..."
            End
        End If
    End If
    
    With PropBag
        'Exe info
        lblText.Caption = .ReadProperty("Text")
        Set imgPic.Picture = .ReadProperty("Picture")
        Me.Caption = .ReadProperty("Caption")
        lblHeader.Caption = .ReadProperty("Header")
    End With
    With frmAbout
        .lblInfo(0).Caption = App.EXEName & ".exe"
        .lblInfo(1).Caption = PropBag.ReadProperty("LOF") & " bytes"
        .lblInfo(2).Caption = PropBag.ReadProperty("BeginPos") & " bytes"
        .lblInfo(3).Caption = (PropBag.ReadProperty("LOF") - PropBag.ReadProperty("BeginPos")) & " bytes"
        .lblInfo(4).Caption = Format((PropBag.ReadProperty("BeginPos") / PropBag.ReadProperty("LOF")) * 100, "0.00") & " %"
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMail.ForeColor = &H0
End Sub

Private Sub lblmail_Click()
    Shell "start mailto:Xbrain3000@hotmail.com"
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMail.ForeColor = &HFF0000
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    On Error Resume Next
    Unload Me
    Unload frmAbout
    End
End Sub
