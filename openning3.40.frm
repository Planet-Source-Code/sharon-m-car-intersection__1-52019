VERSION 5.00
Begin VB.Form frmOpen 
   Caption         =   "Choose Your Cars"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   8310
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      Picture         =   "openning3.40.frx":0000
      TabIndex        =   39
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   7
      Left            =   6960
      TabIndex        =   35
      Top             =   6360
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   6
         Left            =   2640
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: îùôçúé"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   7
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   7
         Left            =   240
         Picture         =   "openning3.40.frx":2417
         Top             =   360
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   6
      Left            =   6960
      TabIndex        =   31
      Top             =   4320
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: îùôçúé"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   6
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   6
         Left            =   240
         Picture         =   "openning3.40.frx":2C5D
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   5
      Left            =   6960
      TabIndex        =   27
      Top             =   2280
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú: ÷î""ù"
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   30
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: îùôçúé"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   5
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   5
         Left            =   240
         Picture         =   "openning3.40.frx":33C2
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   4
      Left            =   6960
      TabIndex        =   24
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú: ÷î""ù"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   40
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: îùôçúé"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   4
         Left            =   240
         Picture         =   "openning3.40.frx":3C69
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   6360
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   8
         Left            =   2640
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: îùàéú"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   3
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   495
         Index           =   3
         Left            =   240
         Picture         =   "openning3.40.frx":4542
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   4320
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: ñôåøè"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   2
         Left            =   240
         Picture         =   "openning3.40.frx":5151
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: ñôåøè"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   1
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   1
         Left            =   240
         Picture         =   "openning3.40.frx":5968
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame frmCar 
      Height          =   1935
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "ùí äøëá: "
         Height          =   375
         Index           =   0
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "ñåâ äøëá: ñôåøè"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "îäéøåú:  ÷î""ù"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgFireCar 
         Height          =   1170
         Index           =   0
         Left            =   240
         Picture         =   "openning3.40.frx":62C8
         Top             =   240
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public boolDevicesSound As Boolean

' Detect if computer has sound card that plays wave audio
#If Win32 Then
    Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
#ElseIf Win16 Then
    Private Declare Function waveOutGetNumDevs Lib "mmsystem" () As Integer
#End If



Private Sub cmdPlay_Click()
    Dim intCarNum As Integer
    Dim intCarCount As Integer
    
    
    intCarCount = 0
    boolDevicesSound = True
    
    'Checking that each car has a name
    For intCarNum = 1 To 8
    
        If txtName(intCarNum).Text = Empty Then
            MsgBox "You must provide a name to each car", vbExclamation, "Attention!"
            txtName(intCarNum).SetFocus
            Exit For
        Else
            MyCars(intCarNum).CarName = txtName(intCarNum).Text
            intCarCount = intCarCount + 1
        End If
    Next intCarNum
    
    'If all the cars have names, start the game
    If intCarCount = 8 Then
        Unload Me
        Load frmIntersection
    End If
    
    
    '============================================
    ' checking if there is a "save file"
    '============================================
    
    filepath = App.Path & "\intersave.txt"
    On Error Resume Next
    fLen = Len(Dir$(filepath))
    If Err Or fLen = 0 Then
        'file doesnt exist
        frmIntersection.cmdLoad.Enabled = False     ' disabeling the "load" button
        mdiIntersection.Load.Enabled = False
    Else
        ' file exist
        Kill App.Path & "\intersave.txt"            'delete "save file"
        frmIntersection.cmdLoad.Enabled = False     ' disabeling the "load" button
        mdiIntersection.Load.Enabled = False
    End If
    
    
    '=======================================================
    'Detect if computer has sound card that plays wave audio
    '=======================================================
    
    #If Win32 Then
        Dim i As Long
    #ElseIf Win16 Then
        Dim i As Integer
    #End If
        
        i = waveOutGetNumDevs()
        If i < 1 Then         ' There is no sound device.
            boolDevicesSound = False
            MsgBox "Can't play wave data"
        Else
            boolDevicesSound = True
        End If
     
End Sub

Private Sub Form_Load()
   Dim intcounter As Integer
    
    For intcounter = 1 To 8     'Set the aligment to the left
        txtName(intcounter).Alignment = 0
    Next intcounter
     
'--------------------------------------------------------------
    ' Set the vehicles properties
'---------------------------------------------------------------
    Vehicle(1).Type = "Sport Car"
    Vehicle(1).Speed = 45
    Vehicle(1).FullTank = 25000
    Vehicle(1).KmPerMove = 100
    
    Vehicle(2).Type = "Normal Car"
    Vehicle(2).Speed = 45
    Vehicle(2).FullTank = 20000
    Vehicle(2).KmPerMove = 70
    
    Vehicle(3).Type = "Truck"
    Vehicle(3).Speed = 40
    Vehicle(3).FullTank = 30000
    Vehicle(3).KmPerMove = 70
    
    Vehicle(4).Type = "Imergency Car"
    Vehicle(4).Speed = 70
    Vehicle(4).FullTank = 20000
    Vehicle(4).KmPerMove = 100
    
'-----------------------------------------------------------------
'   Giving defult names to the cars
'-----------------------------------------------------------------
    txtName(BlueCar).Text = "Baby Blue"
    txtName(FireCar).Text = "Hot Fire"
    txtName(RedCar).Text = "Cute Red"
    txtName(WhiteCar).Text = "Funky White"
    txtName(BlackCar).Text = "Dark Black"
    txtName(GreenCar).Text = "Big Green"
    txtName(YellowCar).Text = "Cool Yellow"
    txtName(TruckCar).Text = "Super Truck"
    
 '--------------------------------------------------------------------
 '  Show the cars speed on screen
 '--------------------------------------------------------------------
    lblSpeed(BlueCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrBlueCar.Interval) & " ÷î""ù"
    lblSpeed(FireCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrFireCar.Interval) & " ÷î""ù"
    lblSpeed(RedCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrRedCar.Interval) & " ÷î""ù"
    lblSpeed(WhiteCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrWhiteCar.Interval) & " ÷î""ù"
    lblSpeed(BlackCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrBlackCar.Interval) & " ÷î""ù"
    lblSpeed(GreenCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrGreenCar.Interval) & " ÷î""ù"
    lblSpeed(YellowCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrYellowCar.Interval) & " ÷î""ù"
    lblSpeed(TruckCar).Caption = "îäéøåú : " & CInt(1000 / frmIntersection.tmrTruckCar.Interval) & " ÷î""ù"
    
End Sub



Private Sub txtName_GotFocus(Index As Integer)
    
    txtName(BlueCar).SelStart = 0
    txtName(BlueCar).SelLength = Len(txtName(BlueCar).Text)

    txtName(RedCar).SelStart = 0
    txtName(RedCar).SelLength = Len(txtName(RedCar).Text)

    txtName(WhiteCar).SelStart = 0
    txtName(WhiteCar).SelLength = Len(txtName(WhiteCar).Text)

    txtName(BlackCar).SelStart = 0
    txtName(BlackCar).SelLength = Len(txtName(BlackCar).Text)

    txtName(GreenCar).SelStart = 0
    txtName(GreenCar).SelLength = Len(txtName(GreenCar).Text)

    txtName(YellowCar).SelStart = 0
    txtName(YellowCar).SelLength = Len(txtName(YellowCar).Text)

    txtName(FireCar).SelStart = 0
    txtName(FireCar).SelLength = Len(txtName(FireCar).Text)
    
    txtName(TruckCar).SelStart = 0
    txtName(TruckCar).SelLength = Len(txtName(TruckCar).Text)
    
   
End Sub

