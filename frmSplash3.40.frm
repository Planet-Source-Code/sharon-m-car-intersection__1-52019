VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash3.40.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9120
      Begin VB.Timer tmrSplash 
         Interval        =   5000
         Left            =   8160
         Top             =   4680
      End
      Begin VB.Label lblOurNames 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   6720
         TabIndex        =   4
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "The Intersection"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   6495
      End
      Begin VB.Image imgLogo 
         Height          =   2625
         Left            =   240
         Picture         =   "frmSplash3.40.frx":000C
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Version  3.40"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   6720
         TabIndex        =   1
         Top             =   3360
         Width           =   1470
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Arik And Sharon Presents:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   705
         Width           =   4560
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Load frmOpen
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblOurNames.Caption = "By :" & vbCrLf & "Arik Havar And" & vbCrLf & "Sharon Melamed"
End Sub

Private Sub Frame1_Click()
    Unload Me
    Load frmOpen
End Sub



Private Sub tmrSplash_Timer()
    Unload Me
    Load frmOpen
End Sub
