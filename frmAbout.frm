VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4200
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2898.915
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtThanks 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   1050
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmAbout.frx":0000
      Top             =   1290
      Width           =   3885
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   1020
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   750
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Additional Info: ..."
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   255
      TabIndex        =   2
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strApplication   As String   'Variable to hold the application name
Public strVersion       As String   'Variable to hold the version info
Public strAuthor        As String   'Variable to hold the author's name
Public strThanks        As String   'Variable to hold the Special Thanks info
Public strAddInfo       As String   'Variable to hold the Addition info
Public picLogo          As Picture  'Variable to hold the logo to use

Private Sub cmdOK_Click()
  Unload Me     'Unload the about box
End Sub

Private Sub Form_Activate()
    lblTitle.Caption = strApplication   'Set the title
    Me.Caption = "About: " & strApplication 'Set the caption of the form
    lblVersion.Caption = "Version: " & strVersion   'Set the version
    lblAuthor.Caption = "By: " & strAuthor  'Set the Author
    txtThanks.Text = strThanks          'Add in the Special Thanks
    txtThanks.Locked = True             'Lock the textbox so that people can still scroll if necessary but can't change
    lblDisclaimer.Caption = strAddInfo  'Set the extra information
    Set picIcon.Picture = picLogo       'Set the picture to use as the logo
    Set Me.Icon = picLogo               'Set the form's picture as well
End Sub
