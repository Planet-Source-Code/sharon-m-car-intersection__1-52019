VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Cars Status"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Back To Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6360
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame fraCar6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgGreenCar 
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul6 
         BorderColor     =   &H80000000&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel6 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM6 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed6 
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusGreen 
         Height          =   375
         Left            =   1680
         Picture         =   "frmStatus3.40.frx":0000
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame fraCar5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgBlackCar 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul5 
         BorderColor     =   &H80000000&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel5 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM5 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed5 
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusBlack 
         Height          =   345
         Left            =   1800
         Picture         =   "frmStatus3.40.frx":10AA
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame fraCar3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgRedCar 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul3 
         BorderColor     =   &H80000004&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel3 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM3 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblStatusSpeed3 
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusRed 
         Height          =   345
         Left            =   1800
         Picture         =   "frmStatus3.40.frx":1DDC
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame fraCar1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgBlueCar 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul1 
         BorderColor     =   &H80000004&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed1 
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusBlue 
         Height          =   345
         Left            =   1680
         Picture         =   "frmStatus3.40.frx":2B0E
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraCar8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgTruckCar 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul8 
         BorderColor     =   &H80000000&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel8 
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM8 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed8 
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusTruck 
         Height          =   345
         Left            =   1560
         Picture         =   "frmStatus3.40.frx":39B0
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame fraCar7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgYellowCar 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul7 
         BorderColor     =   &H80000000&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel7 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM7 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed7 
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusYellow 
         Height          =   315
         Left            =   1800
         Picture         =   "frmStatus3.40.frx":4CA2
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame fraCar4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgWhiteCar 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul4 
         BorderColor     =   &H80000004&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel4 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM4 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed4 
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusWhite 
         Height          =   315
         Left            =   1680
         Picture         =   "frmStatus3.40.frx":58B4
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraCar2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin MSComctlLib.ProgressBar prgFireCar 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape shpNeedFeul2 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H80000004&
         FillColor       =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblStatusFuel2 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblStatusKM2 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblStatusSpeed2 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgStatusFire 
         Height          =   255
         Left            =   1680
         Picture         =   "frmStatus3.40.frx":6616
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub cmdGoBack_Click()
    frmIntersection.WindowState = 2
    frmIntersection.Show
    Unload Me
End Sub

Private Sub Form_Load()

    'Set the cars names on the frame
    
    fraCar1.Caption = MyCars(BlueCar).CarName
    fraCar2.Caption = MyCars(FireCar).CarName
    fraCar3.Caption = MyCars(RedCar).CarName
    fraCar4.Caption = MyCars(WhiteCar).CarName
    fraCar5.Caption = MyCars(BlackCar).CarName
    fraCar6.Caption = MyCars(GreenCar).CarName
    fraCar7.Caption = MyCars(YellowCar).CarName
    fraCar8.Caption = MyCars(TruckCar).CarName
    
    'Set the max value of the scrolls bar in Status form
         
    frmStatus.prgBlueCar.Max = 20000
    frmStatus.prgFireCar.Max = 25000
    frmStatus.prgRedCar.Max = 20000
    frmStatus.prgWhiteCar.Max = 25000
    frmStatus.prgBlackCar.Max = 20000
    frmStatus.prgGreenCar.Max = 20000
    frmStatus.prgYellowCar.Max = 25000
    frmStatus.prgTruckCar.Max = 30000
End Sub

Public Sub ShowStatus() 'Updating the cars properties
    
    'Blue Car
    frmStatus.prgBlueCar.Value = MyCars(BlueCar).Fuel.FuelForNow
    frmStatus.lblStatusKM1.Caption = "Km : " & MyCars(BlueCar).KM
    frmStatus.lblStatusFuel1.Caption = "Fuel Left : " & MyCars(BlueCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed1.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrBlueCar.Interval) & " ÷î""ù"
    If MyCars(BlueCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul1.FillColor = vbRed
    Else
        shpNeedFeul1.FillColor = vbGreen
    End If
    
    'Fire car
    frmStatus.prgFireCar.Value = MyCars(FireCar).Fuel.FuelForNow
    frmStatus.lblStatusKM2.Caption = "Km : " & MyCars(FireCar).KM
    frmStatus.lblStatusFuel2.Caption = "Fuel Left : " & MyCars(FireCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed2.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrFireCar.Interval) & " ÷î""ù"
    If MyCars(FireCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul2.FillColor = vbRed
    Else
        shpNeedFeul2.FillColor = vbGreen
    End If
         
    'Red Car
    frmStatus.prgRedCar.Value = MyCars(RedCar).Fuel.FuelForNow
    frmStatus.lblStatusKM3.Caption = "Km : " & MyCars(RedCar).KM
    frmStatus.lblStatusFuel3.Caption = "Fuel Left : " & MyCars(RedCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed3.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrRedCar.Interval) & " ÷î""ù"
    If MyCars(RedCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul3.FillColor = vbRed
    Else
        shpNeedFeul3.FillColor = vbGreen
    End If
    
    'White Car
    frmStatus.prgWhiteCar.Value = MyCars(WhiteCar).Fuel.FuelForNow
    frmStatus.lblStatusKM4.Caption = "Km : " & MyCars(WhiteCar).KM
    frmStatus.lblStatusFuel4.Caption = "Fuel Left : " & MyCars(WhiteCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed4.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrWhiteCar.Interval) & " ÷î""ù"
    If MyCars(WhiteCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul4.FillColor = vbRed
    Else
        shpNeedFeul4.FillColor = vbGreen
    End If
    
    'Black Car
    frmStatus.prgBlackCar.Value = MyCars(BlackCar).Fuel.FuelForNow
    frmStatus.lblStatusKM5.Caption = "Km : " & MyCars(BlackCar).KM
    frmStatus.lblStatusFuel5.Caption = "Fuel Left : " & MyCars(BlackCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed5.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrBlackCar.Interval) & " ÷î""ù"
    If MyCars(BlackCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul5.FillColor = vbRed
    Else
        shpNeedFeul5.FillColor = vbGreen
    End If
    
    'Green car
    frmStatus.prgGreenCar.Value = MyCars(GreenCar).Fuel.FuelForNow
    frmStatus.lblStatusKM6.Caption = "Km : " & MyCars(GreenCar).KM
    frmStatus.lblStatusFuel6.Caption = "Fuel Left : " & MyCars(GreenCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed6.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrGreenCar.Interval) & " ÷î""ù"
    If MyCars(GreenCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul6.FillColor = vbRed
    Else
        shpNeedFeul6.FillColor = vbGreen
    End If
    
    'Yellow Car
    frmStatus.prgYellowCar.Value = MyCars(YellowCar).Fuel.FuelForNow
    frmStatus.lblStatusKM7.Caption = "Km : " & MyCars(YellowCar).KM
    frmStatus.lblStatusFuel7.Caption = "Fuel Left : " & MyCars(YellowCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed7.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrYellowCar.Interval) & " ÷î""ù"
    If MyCars(YellowCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul7.FillColor = vbRed
    Else
        shpNeedFeul7.FillColor = vbGreen
    End If
    
    'Truck
    frmStatus.prgTruckCar.Value = MyCars(TruckCar).Fuel.FuelForNow
    frmStatus.lblStatusKM8.Caption = "Km : " & MyCars(TruckCar).KM
    frmStatus.lblStatusFuel8.Caption = "Fuel Left : " & MyCars(TruckCar).Fuel.FuelForNow
    frmStatus.lblStatusSpeed8.Caption = "Speed : " & CInt(1000 / frmIntersection.tmrTruckCar.Interval) & " ÷î""ù"
    If MyCars(TruckCar).Fuel.FuelForNow < intNeedFuel Then
        shpNeedFeul8.FillColor = vbRed
    Else
        shpNeedFeul8.FillColor = vbGreen
    End If
    
    frmStatus.Show
End Sub
