VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmIntersection 
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmVersion 3.40.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   13875
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9720
      Top             =   6600
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "Go To Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowStatus 
      Caption         =   "Show Status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Timer tmrCallPolice 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   7080
      Top             =   7800
   End
   Begin VB.Timer tmrPolice 
      Interval        =   24
      Left            =   8520
      Top             =   7800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton opbLightChange 
      Caption         =   "Auto"
      Height          =   255
      Index           =   1
      Left            =   9840
      TabIndex        =   9
      Top             =   5520
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton opbLightChange 
      Caption         =   "Manual"
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   8
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frmChangeLight 
      Height          =   1335
      Left            =   9720
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.PictureBox picIntersection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9570
      Left            =   0
      Picture         =   "frmVersion 3.40.frx":2C8E
      ScaleHeight     =   9540
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin MCI.MMControl MMcontrol 
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   661
         _Version        =   393216
         PlayEnabled     =   -1  'True
         PlayVisible     =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.PictureBox picPoliceStation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7080
         Picture         =   "frmVersion 3.40.frx":11D56
         ScaleHeight     =   345
         ScaleWidth      =   1785
         TabIndex        =   18
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Timer tmrTruckCar 
         Left            =   2400
         Top             =   8640
      End
      Begin VB.Timer tmrYellowCar 
         Left            =   1920
         Top             =   8640
      End
      Begin VB.Timer tmrGreenCar 
         Left            =   1440
         Top             =   8640
      End
      Begin VB.Timer tmrBlackCar 
         Left            =   960
         Top             =   8640
      End
      Begin VB.Timer tmrWhiteCar 
         Left            =   2280
         Top             =   9120
      End
      Begin VB.PictureBox picSonol 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8880
         Picture         =   "frmVersion 3.40.frx":12576
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picAlon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         Picture         =   "frmVersion 3.40.frx":12A44
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.Timer tmrRedCar 
         Left            =   1800
         Top             =   9120
      End
      Begin VB.Timer tmrYellowLight 
         Interval        =   1
         Left            =   480
         Top             =   0
      End
      Begin VB.Timer tmrFireCar 
         Left            =   1320
         Top             =   9120
      End
      Begin VB.Timer tmrFuel2 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   960
         Top             =   1920
      End
      Begin VB.Timer tmrFuel1 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   6960
         Top             =   1080
      End
      Begin VB.Timer tmrLableFade 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   7680
         Top             =   600
      End
      Begin VB.Timer tmrLights 
         Left            =   840
         Top             =   9120
      End
      Begin VB.Timer tmrBlueCar 
         Left            =   360
         Top             =   9120
      End
      Begin VB.Shape shpBlueLight 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   8160
         Shape           =   3  'Circle
         Top             =   8520
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpRedLight 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   8520
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgLight5 
         Height          =   420
         Left            =   3000
         Picture         =   "frmVersion 3.40.frx":12EF0
         Top             =   5400
         Width           =   270
      End
      Begin VB.Image imgLight4 
         Height          =   420
         Left            =   3000
         Picture         =   "frmVersion 3.40.frx":13552
         Top             =   4800
         Width           =   270
      End
      Begin VB.Image imgLight22 
         Height          =   375
         Left            =   4680
         Picture         =   "frmVersion 3.40.frx":13BB4
         Top             =   6000
         Width           =   225
      End
      Begin VB.Image imgLight23 
         Height          =   375
         Left            =   5280
         Picture         =   "frmVersion 3.40.frx":140A6
         Top             =   6000
         Width           =   225
      End
      Begin VB.Image imgLight9 
         Height          =   375
         Left            =   6000
         Picture         =   "frmVersion 3.40.frx":14598
         Top             =   4080
         Width           =   225
      End
      Begin VB.Image imgLight8 
         Height          =   375
         Left            =   6000
         Picture         =   "frmVersion 3.40.frx":14A8A
         Top             =   3360
         Width           =   225
      End
      Begin VB.Image imgLight15 
         Height          =   375
         Left            =   4080
         Picture         =   "frmVersion 3.40.frx":14F7C
         Top             =   2880
         Width           =   225
      End
      Begin VB.Image imgLight14 
         Height          =   375
         Left            =   3480
         Picture         =   "frmVersion 3.40.frx":1546E
         Top             =   2880
         Width           =   225
      End
      Begin VB.Image imgPolice 
         Height          =   405
         Left            =   7560
         Picture         =   "frmVersion 3.40.frx":15960
         Top             =   7320
         Width           =   900
      End
      Begin VB.Label lblCarFuelName1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblCarFuelName2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image imgGreenCar 
         Height          =   405
         Left            =   0
         Picture         =   "frmVersion 3.40.frx":16C9E
         Top             =   9000
         Width           =   900
      End
      Begin VB.Image imgYellowCar 
         Height          =   855
         Left            =   1440
         Picture         =   "frmVersion 3.40.frx":172BC
         Top             =   7560
         Width           =   375
      End
      Begin VB.Image imgTruckCar 
         Height          =   1035
         Left            =   960
         Picture         =   "frmVersion 3.40.frx":183EA
         Top             =   7440
         Width           =   345
      End
      Begin VB.Image imgBlackCar 
         Height          =   405
         Left            =   840
         Picture         =   "frmVersion 3.40.frx":19794
         Top             =   6960
         Width           =   855
      End
      Begin VB.Image imgWhiteCar 
         Height          =   360
         Left            =   0
         Picture         =   "frmVersion 3.40.frx":1A9FA
         Top             =   8520
         Width           =   885
      End
      Begin VB.Image imgRedCar 
         Height          =   375
         Left            =   0
         Picture         =   "frmVersion 3.40.frx":1BB1C
         Top             =   8040
         Width           =   780
      End
      Begin VB.Image imgFireCar 
         Height          =   255
         Left            =   0
         Picture         =   "frmVersion 3.40.frx":1CA9A
         Top             =   7560
         Width           =   840
      End
      Begin VB.Label lblCarKM 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblCarFuel 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblCarSpeed 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblCarName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image imgBlueCar 
         Height          =   345
         Left            =   0
         Picture         =   "frmVersion 3.40.frx":1D604
         Top             =   6960
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000004&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   7080
         Top             =   7200
         Width           =   1815
      End
   End
   Begin VB.Label lblTimeSec 
      Height          =   255
      Left            =   10800
      TabIndex        =   25
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   ":"
      Height          =   255
      Left            =   10680
      TabIndex        =   24
      Top             =   6120
      Width           =   195
   End
   Begin VB.Label lblTimeMin 
      Height          =   255
      Left            =   10320
      TabIndex        =   23
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   ":"
      Height          =   255
      Left            =   10200
      TabIndex        =   22
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label lblTimeHour 
      Height          =   255
      Left            =   9840
      TabIndex        =   21
      Top             =   6120
      Width           =   255
   End
End
Attribute VB_Name = "frmIntersection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
 
 Dim Lights(8) As Integer
 Dim FirstInterval(8) As Boolean
 Public OpenFile As Integer
 Public LightManual As Boolean
 Public FireCarsquare As Integer
 Public BlueCarsquare As Integer
 Public PoliceCarsquare As Integer
 Public RedCarsquare As Integer
 Public WhiteCarsquare As Integer
 Public BlackCarsquare As Integer
 Public GreenCarsquare As Integer
 Public TruckCarsquare As Integer
 Public YellowCarsquare As Integer
  
 
Private Sub cmdChange_Click()   'Changing the traffic lights manualy

    tmrLights.Enabled = True
    LightManual = True
    
End Sub

Public Sub cmdCode_Click()
    Unload frmIntersection
    Unload frmStatus
    Unload mdiIntersection
    
End Sub

Public Sub cmdExit_Click()      'Exit the game without saving
    
    Dim MyAnswer
    Dim j As Integer
     
    MyAnswer = MsgBox("Are you sure you want to quit without saving?", vbYesNo + vbQuestion, "Exit")
    If MyAnswer = vbYes Then
        Unload Me
        Unload frmStatus
        Unload frmOpen
        Load Me
        For j = 1 To 8
            MyCars(j).Fuel.FuelForNow = 0
            MyCars(j).KM = 0
        Next j
                
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
        frmIntersection.cmdLoad.Enabled = True     ' enabling the "load" button
        mdiIntersection.Load.Enabled = True
    End If

    
End Sub

Public Sub cmdLoad_Click()      'Loading the last saving of the game
    
    Dim MyAnswer
    Dim reclen As Integer
    Dim intIndex As Integer
    Dim i As Integer
    Dim intSquares As Integer
    
    
    MyAnswer = MsgBox("Are you sure you want to load an old game?", vbYesNo + vbQuestion, "Load")
    If MyAnswer = vbYes Then     'If want to load
        OpenFile = FreeFile
        reclen = Len(SavingDetails)
        Open App.Path & "/InterSave.txt" For Random As #OpenFile Len = reclen
                Get #OpenFile, , SavingDetails
        Close #OpenFile
                
        Call mdlSaveNLoad.fnLoadingDetails(BlueCar, imgBlueCar, "Blue", tmrBlackCar)
        Call mdlSaveNLoad.fnLoadingDetails(RedCar, imgRedCar, "Red", tmrRedCar)
        Call mdlSaveNLoad.fnLoadingDetails(FireCar, imgFireCar, "Fire", tmrFireCar)
        Call mdlSaveNLoad.fnLoadingDetails(GreenCar, imgGreenCar, "Green", tmrGreenCar)
        Call mdlSaveNLoad.fnLoadingDetails(WhiteCar, imgWhiteCar, "White", tmrWhiteCar)
        Call mdlSaveNLoad.fnLoadingDetails(BlackCar, imgBlackCar, "black", tmrBlackCar)
        Call mdlSaveNLoad.fnLoadingDetails(YellowCar, imgYellowCar, "Yellow", tmrYellowCar)
        Call mdlSaveNLoad.fnLoadingDetails(TruckCar, imgTruckCar, "Truck", tmrTruckCar)
        Call mdlSaveNLoad.fnLoadingDetails(Police, imgPolice, "Imergency", tmrCallPolice)
        
        
        For i = 1 To 12
            LineBusy(i) = SavingDetails.SaveLineBusy(i)
            CarsInLine(i) = SavingDetails.SaveCarsInLine(i)
        Next i
        
        For intSquares = 1 To SquaresNum
            LinesSquares4(intSquares) = SavingDetails.linesSquares(1, intSquares)
            LinesSquares5(intSquares) = SavingDetails.linesSquares(2, intSquares)
            LinesSquares6(intSquares) = SavingDetails.linesSquares(3, intSquares)
            LinesSquares7(intSquares) = SavingDetails.linesSquares(4, intSquares)
            LinesSquares8(intSquares) = SavingDetails.linesSquares(5, intSquares)
            LinesSquares9(intSquares) = SavingDetails.linesSquares(6, intSquares)
            LinesSquares13(intSquares) = SavingDetails.linesSquares(7, intSquares)
            LinesSquares14(intSquares) = SavingDetails.linesSquares(8, intSquares)
            LinesSquares15(intSquares) = SavingDetails.linesSquares(9, intSquares)
            LinesSquares22(intSquares) = SavingDetails.linesSquares(10, intSquares)
            LinesSquares23(intSquares) = SavingDetails.linesSquares(11, intSquares)
            LinesSquares24(intSquares) = SavingDetails.linesSquares(12, intSquares)
        Next intSquares
        
        If SavingDetails.Fuel1IsWorking = True Then
            tmrFuel1.Enabled = True
            boolFuel1Full = True
        End If
        
        If SavingDetails.Fuel2IsWorking = True Then
            tmrFuel2.Enabled = True
            boolFuel2Full = True
        End If
        
        If intCarsInRest = 9 Then
            Form_Load
            cmdNewGame_Click
        End If
        
        tmrCallPolice.Enabled = True
        CarWay(Police) = "Start"
        tmrTimer.Enabled = True
        tmrYellowLight.Enabled = True
        
    End If
    
    cmdLoad.Enabled = True    ' enabling the "load" button
    mdiIntersection.Load.Enabled = True
    
End Sub

Public Sub cmdNewGame_Click()       'Start a new game

    Dim MyAnswer
    
    MySec = 0
    MyMin = 0
    MyHour = 0
    tmrTimer.Enabled = True
    
    'Set the Start position of the cars in the cars square's to one
    
    boolFuel1Full = False
    boolFuel2Full = False
    FireCarsquare = 1
    BlueCarsquare = 1
    PoliceCarsquare = 1
    RedCarsquare = 1
    BlackCarsquare = 1
    WhiteCarsquare = 1
    GreenCarsquare = 1
    YellowCarsquare = 1
    TruckCarsquare = 1
      
    intMyLight = 0
    boolYellow = True
    'How many cars are in the game
    CarNumber = 9
        
    'If it is a total new game or clicking "NEW GAME" during the game
    If MyCars(BlueCar).KM = 0 Then
        tmrTruckCar.Enabled = True
        tmrLights.Enabled = True
        tmrYellowLight.Enabled = True
        Randomize

        CarWay(BlueCar) = "End"
        CarWay(FireCar) = "End"
        CarWay(RedCar) = "End"
        CarWay(WhiteCar) = "End"
        CarWay(BlackCar) = "End"
        CarWay(GreenCar) = "End"
        CarWay(YellowCar) = "End"
        CarWay(TruckCar) = "End"
    Else
        MyAnswer = MsgBox("Are you sure ou want to start a new game and quit this one?", vbYesNo + vbQuestion, "Attention")
        If MyAnswer = vbYes Then
           cmdExit_Click
           cmdNewGame_Click
        End If
    End If
     
     mdlStartInit.NewGameInitialazetion
    
End Sub

Public Sub cmdSave_Click()      'Saving the game
         
    Dim reclen As Integer
    Dim MyAnswer
    Dim intIndex As Integer
    Dim LineCount As Integer
    Dim intCarNum As Integer
    Dim intRow As Integer
    
    MyAnswer = MsgBox("Are you sure you want to save and exit the game?", vbYesNo + vbQuestion, "Save")
    If MyAnswer = vbYes Then
    
    intCarsInRest = 1
    
    'Blue car details
        Call mdlSaveNLoad.fnSavingDetails(BlueCar, imgBlueCar, BlueCarsquare)
        
    'Fire car details
        Call mdlSaveNLoad.fnSavingDetails(FireCar, imgFireCar, FireCarsquare)
        
     'Yellow car details
        Call mdlSaveNLoad.fnSavingDetails(YellowCar, imgYellowCar, YellowCarsquare)
        
    'Red car details
        Call mdlSaveNLoad.fnSavingDetails(RedCar, imgRedCar, RedCarsquare)
        
    'White car details
        Call mdlSaveNLoad.fnSavingDetails(WhiteCar, imgWhiteCar, WhiteCarsquare)
        
    'Black car details
        Call mdlSaveNLoad.fnSavingDetails(BlackCar, imgBlackCar, BlackCarsquare)
        
    'Green car details
        Call mdlSaveNLoad.fnSavingDetails(GreenCar, imgGreenCar, GreenCarsquare)
            
    'Truck car details
        Call mdlSaveNLoad.fnSavingDetails(TruckCar, imgTruckCar, TruckCarsquare)
        
    'Police car details
        Call mdlSaveNLoad.fnSavingDetails(Police, imgPolice, PoliceCarsquare)
        
        If boolFuel1Full = True Then
            SavingDetails.Fuel1IsWorking = True
        Else
            SavingDetails.Fuel1IsWorking = False
        End If
        
        If boolFuel2Full = True Then
            SavingDetails.Fuel2IsWorking = True
        Else
            SavingDetails.Fuel2IsWorking = False
        End If
        
        For LineCount = 1 To 12
            SavingDetails.SaveCarsInLine(LineCount) = CarsInLine(LineCount)
            SavingDetails.SaveLineBusy(LineCount) = LineBusy(LineCount)
        Next LineCount
        
        For intCarNum = 1 To 9
            If boolSaveInRest(intCarNum) = True Then
                intCarsInRest = intCarsInRest + 1
            End If
        Next intCarNum
        
       
        For intRow = 1 To SquaresNum
            SavingDetails.linesSquares(1, intRow) = LinesSquares4(intRow)
            SavingDetails.linesSquares(2, intRow) = LinesSquares5(intRow)
            SavingDetails.linesSquares(3, intRow) = LinesSquares6(intRow)
            SavingDetails.linesSquares(4, intRow) = LinesSquares7(intRow)
            SavingDetails.linesSquares(5, intRow) = LinesSquares8(intRow)
            SavingDetails.linesSquares(6, intRow) = LinesSquares9(intRow)
            SavingDetails.linesSquares(7, intRow) = LinesSquares13(intRow)
            SavingDetails.linesSquares(8, intRow) = LinesSquares14(intRow)
            SavingDetails.linesSquares(9, intRow) = LinesSquares15(intRow)
            SavingDetails.linesSquares(10, intRow) = LinesSquares22(intRow)
            SavingDetails.linesSquares(11, intRow) = LinesSquares23(intRow)
            SavingDetails.linesSquares(12, intRow) = LinesSquares24(intRow)
        Next intRow
       
        
        reclen = Len(SavingDetails)
        OpenFile = FreeFile
        Open App.Path & "\InterSave.txt" For Random As #OpenFile Len = reclen
            Put #OpenFile, , SavingDetails
        Close #OpenFile
        
        For intCarNum = 1 To 8
            MyCars(intCarNum).Fuel.FuelForNow = 0
            MyCars(intCarNum).KM = 0
        Next intCarNum
        Unload Me
        Load Me
        mdiIntersection.Load.Enabled = True
    End If
    
End Sub

Private Sub cmdShowStatus_Click()
    Load frmStatus
    Call frmStatus.ShowStatus
End Sub


Public Sub Form_Load()

     Randomize
    tmrBlueCar.Enabled = False
    tmrFireCar.Enabled = False
    tmrRedCar.Enabled = False
    tmrWhiteCar.Enabled = False
    tmrBlackCar.Enabled = False
    tmrGreenCar.Enabled = False
    tmrYellowCar.Enabled = False
    tmrTruckCar.Enabled = False
    tmrPolice.Enabled = False
    tmrLights.Enabled = False
    tmrYellowLight.Enabled = False
    lblCarName.Visible = False
    lblCarFuel.Visible = False
    lblCarKM.Visible = False
    lblCarSpeed.Visible = False
    cmdChange.Enabled = False
    FuelPerTimer = 10
    intNeedFuel = 7000
    intHaveToSlow = 25
    intAdd2Car = 100
    

   
'------------------------------------------------------------------
    'Intervals Settings
'------------------------------------------------------------------

    tmrFireCar.Interval = 30
    tmrYellowCar.Interval = 30
    tmrWhiteCar.Interval = 30
    tmrBlueCar.Interval = 50    'Blue car Interval
    tmrRedCar.Interval = 50
    tmrGreenCar.Interval = 50
    tmrBlackCar.Interval = 50
    tmrTruckCar.Interval = 60
    tmrLights.Interval = 3000   'Traffic lights interval
    
    
    Dim i As Integer
    For i = 1 To 8
        FirstInterval(i) = True
    Next i
    
    lblTimeSec.Caption = "00"
    lblTimeMin.Caption = "00"
    lblTimeHour.Caption = "00"
    mdlStartInit.FormInitialazetion
    
    
End Sub



Private Sub imgBlueCar_Click()
    
    'Will show the blue car details
    lblCarName.Visible = CarClick(BlueCar, BlueCarsquare, tmrBlueCar)

End Sub

Private Sub imgFireCar_Click()

    'Will show the fire car details
    lblCarName.Visible = CarClick(FireCar, FireCarsquare, tmrFireCar)

End Sub

Private Sub imgPolice_Click()

    lblCarName.Caption = "Police"
    tmrLableFade.Enabled = True
    lblCarName.Visible = True
    
End Sub

Private Sub imgTruckCar_Click()

    'Will show the truck car details
    lblCarName.Visible = CarClick(TruckCar, TruckCarsquare, tmrTruckCar)

End Sub
Private Sub imgYellowCar_Click()

    'Will show the Yellow car details
    lblCarName.Visible = CarClick(YellowCar, YellowCarsquare, tmrYellowCar)

End Sub

Private Sub imgRedCar_Click()

    'Will show the Red car details
    lblCarName.Visible = CarClick(RedCar, RedCarsquare, tmrRedCar)
    
End Sub

Private Sub imgBlackCar_Click()

    'Will show the Black car details
    lblCarName.Visible = CarClick(BlackCar, BlackCarsquare, tmrBlackCar)
    
End Sub
Private Sub imgGreenCar_Click()

    'Will show the green car details
    lblCarName.Visible = CarClick(GreenCar, GreenCarsquare, tmrGreenCar)
    
End Sub

Private Sub imgWhiteCar_Click()

    'Will show the White car details
    lblCarName.Visible = CarClick(WhiteCar, WhiteCarsquare, tmrWhiteCar)

End Sub



Private Sub opbLightChange_Click(Index As Integer)

    Select Case Index
    Case 1      'Change traffic light automatickly
        tmrLights.Enabled = True
        LightManual = False
        cmdChange.Enabled = False
    Case 0      'Change traffic light manualy
        cmdChange.Enabled = True
        tmrLights.Enabled = False
        LightManual = True
    End Select
    
End Sub

Private Sub tmrBlackCar_Timer()
       
    If CarWay(BlackCar) = "Start" Then    'if the car start a new line
        Call mdlSetCases.StartDriving(BlackCar, imgBlackCar, "Black")
    End If
'-------------------------------------------------------------------------------------------------------
                'tart driving
'-------------------------------------------------------------------------------------------------------
    If CarWay(BlackCar) = "Move" Then
         Call mdlMoveInLines.MoveInLine(BlackCar, imgBlackCar, BlackCarsquare, Normal, tmrBlackCar)
    End If
    
    If FirstInterval(BlackCar) = True Then  'Enable the next car interval
        tmrGreenCar.Enabled = True
        FirstInterval(BlackCar) = False
    End If

End Sub

Private Sub tmrBlueCar_Timer()
     
    If CarWay(BlueCar) = "Start" Then    'if the car start a new line
       Call StartDriving(BlueCar, imgBlueCar, "Blue")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(BlueCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(BlueCar, imgBlueCar, BlueCarsquare, Normal, tmrBlueCar)
    End If
    
    If FirstInterval(BlueCar) = True Then  'Enable the next car interval
        tmrRedCar.Enabled = True
        FirstInterval(BlueCar) = False
    End If
            
End Sub

Public Sub tmrCallPolice_Timer()
    If tmrPolice.Enabled = True Then
        tmrPolice.Enabled = False
        MMcontrol.Command = "close"   ' closeing mmcontrol ( Police car siren )
    Else
        tmrPolice.Enabled = True
        shpRedLight.Visible = True
    End If
    
End Sub

Private Sub tmrFireCar_Timer()

    If CarWay(FireCar) = "Start" Then     'if the car start a new line
        Call StartDriving(FireCar, imgFireCar, "Fire")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(FireCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(FireCar, imgFireCar, FireCarsquare, Sport, tmrFireCar)
    End If
    
    If FirstInterval(FireCar) = True Then  'Enable the next car interval
        tmrWhiteCar.Enabled = True
        FirstInterval(FireCar) = False
    End If
    
End Sub

Private Sub tmrFuel1_Timer()

    Select Case CarIsFueling1
        
        Case BlueCar
            MyCars(BlueCar).Fuel.FuelForNow = mdlFuel.Fueling1(BlueCar, imgBlueCar, tmrBlueCar)
        
        Case FireCar
            MyCars(FireCar).Fuel.FuelForNow = mdlFuel.Fueling1(FireCar, imgFireCar, tmrFireCar)
        
        Case RedCar
            MyCars(RedCar).Fuel.FuelForNow = mdlFuel.Fueling1(RedCar, imgRedCar, tmrRedCar)
            
        Case GreenCar
            MyCars(GreenCar).Fuel.FuelForNow = mdlFuel.Fueling1(GreenCar, imgGreenCar, tmrGreenCar)
            
        Case BlackCar
            MyCars(BlackCar).Fuel.FuelForNow = mdlFuel.Fueling1(BlackCar, imgBlackCar, tmrBlackCar)
        
        Case WhiteCar
            MyCars(WhiteCar).Fuel.FuelForNow = mdlFuel.Fueling1(WhiteCar, imgWhiteCar, tmrWhiteCar)
            
        Case YellowCar
            MyCars(YellowCar).Fuel.FuelForNow = mdlFuel.Fueling1(YellowCar, imgYellowCar, tmrYellowCar)
        
        Case TruckCar
            MyCars(TruckCar).Fuel.FuelForNow = mdlFuel.Fueling1(TruckCar, imgTruckCar, tmrTruckCar)
            
        End Select
        
        tmrFuel1.Enabled = False
        lblCarFuelName1.Visible = False
        picSonol.Visible = False

End Sub

Private Sub tmrFuel2_Timer()

        Select Case CarIsFueling2
        
        Case BlueCar
            MyCars(BlueCar).Fuel.FuelForNow = mdlFuel.Fueling2(BlueCar, imgBlueCar, tmrBlueCar)
        
        Case FireCar
            MyCars(FireCar).Fuel.FuelForNow = mdlFuel.Fueling2(FireCar, imgFireCar, tmrFireCar)
        
        Case RedCar
            MyCars(RedCar).Fuel.FuelForNow = mdlFuel.Fueling2(RedCar, imgRedCar, tmrRedCar)
            
        Case GreenCar
            MyCars(GreenCar).Fuel.FuelForNow = mdlFuel.Fueling2(GreenCar, imgGreenCar, tmrGreenCar)
        
        Case BlackCar
            MyCars(BlackCar).Fuel.FuelForNow = mdlFuel.Fueling2(BlackCar, imgBlackCar, tmrBlackCar)
            
        Case WhiteCar
            MyCars(WhiteCar).Fuel.FuelForNow = mdlFuel.Fueling2(WhiteCar, imgWhiteCar, tmrWhiteCar)
            
        Case YellowCar
            MyCars(YellowCar).Fuel.FuelForNow = mdlFuel.Fueling2(YellowCar, imgYellowCar, tmrYellowCar)
            
        Case TruckCar
            MyCars(TruckCar).Fuel.FuelForNow = mdlFuel.Fueling2(TruckCar, imgTruckCar, tmrTruckCar)
            
        End Select
        
        tmrFuel2.Enabled = False
        lblCarFuelName2.Visible = False
        picAlon.Visible = False
       
End Sub

Private Sub tmrGreenCar_Timer()
        
    If CarWay(GreenCar) = "Start" Then    'if the car start a new line
       Call StartDriving(GreenCar, imgGreenCar, "Green")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(GreenCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(GreenCar, imgGreenCar, GreenCarsquare, Normal, tmrGreenCar)
    End If
    
    If FirstInterval(GreenCar) = True Then  'Enable the next car interval
        tmrFireCar.Enabled = True
        FirstInterval(GreenCar) = False
    End If
    
End Sub

Private Sub tmrLableFade_Timer()
    lblCarName.Visible = False
    lblCarFuel.Visible = False
    lblCarKM.Visible = False
    lblCarSpeed.Visible = False
End Sub

Public Sub tmrLights_Timer()

    If boolYellow = True Then
        tmrYellowLight.Enabled = True
        tmrLights.Enabled = False
    Else
      
        Select Case intMyLight
            Case 1
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor5 = "GREEN"
                strLightColor4 = "GREEN"
                Call mdlCheckEntery.MoveCars(5, 4)
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 5
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor9 = "GREEN"
                strLightColor4 = "GREEN"
                Call mdlCheckEntery.MoveCars(9, 4)
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 6
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor22 = "GREEN"
                strLightColor8 = "GREEN"
                Call mdlCheckEntery.MoveCars(8, 22)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 3
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor9 = "GREEN"
                strLightColor8 = "GREEN"
                Call mdlCheckEntery.MoveCars(8, 9)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
           
            Case 4
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor15 = "GREEN"
                strLightColor14 = "GREEN"
                Call mdlCheckEntery.MoveCars(15, 14)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 7
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor15 = "GREEN"
                strLightColor5 = "GREEN"
                Call mdlCheckEntery.MoveCars(15, 5)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor23 = "RED"
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 2
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor22 = "GREEN"
                strLightColor23 = "GREEN"
                Call mdlCheckEntery.MoveCars(22, 23)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor14 = "RED"
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
            Case 8
                imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrgs.bmp")
                strLightColor14 = "GREEN"
                strLightColor23 = "GREEN"
                Call mdlCheckEntery.MoveCars(14, 23)
                strLightColor4 = "RED"
                imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor5 = "RED"
                imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor8 = "RED"
                imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor9 = "RED"
                imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor15 = "RED"
                imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
                strLightColor22 = "RED"
                imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            
        End Select
    End If
    If LightManual = True Then
        tmrLights = False
    End If
    boolYellow = True
    
End Sub

Private Sub tmrPolice_Timer()

Dim length As Long

  If CarWay(Police) = "Start" Then
        CarIntersection(Police) = mdlPolice.PoliceLine
        Call mdlPolice.SetPoliceCar
  End If
    
If CarWay(Police) = "Move" Then
    Call mdlPolice.MovePolice(PoliceCarsquare)
    Call mdlPolice.TernLightsRed
    tmrLights.Enabled = False
    If frmIntersection.shpRedLight.Visible = True Then
        frmIntersection.shpBlueLight.Visible = True
        frmIntersection.shpRedLight.Visible = False
    Else
        frmIntersection.shpBlueLight.Visible = False
        frmIntersection.shpRedLight.Visible = True
    End If
    
    
    ' Play siren while police car move
    
    If frmOpen.boolDevicesSound = True Then
        'Set properties needed by MCI to open ( siren of police car )
        MMcontrol.Notify = False
        MMcontrol.Wait = False
        MMcontrol.Shareable = False
        MMcontrol.FileName = App.Path & "\images\police.wav"
        MMcontrol.Command = "Open"
        length = MMcontrol.length
        MMcontrol.Command = "play"
        If MMcontrol.Position = length Then
            MMcontrol.Command = "prev"
            MMcontrol.Command = "play"
        End If
    Else
        ' close MMcontrol ( police car siren )
        MMcontrol.Command = "Close"
    
    End If
    
End If
     
    If (imgPolice.Top = CarsStartPosition(Police).Top) Then
        frmIntersection.shpRedLight.Visible = False
        frmIntersection.shpBlueLight.Visible = False
    End If
       
End Sub

Private Sub tmrRedCar_Timer()

    If CarWay(RedCar) = "Start" Then    'if the car start a new line
       Call StartDriving(RedCar, imgRedCar, "Red")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(RedCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(RedCar, imgRedCar, RedCarsquare, Normal, tmrRedCar)
    End If
    
    If FirstInterval(RedCar) = True Then  'Enable the next car interval
        tmrBlackCar.Enabled = True
        FirstInterval(RedCar) = False
    End If
            
End Sub

Public Sub tmrTimer_Timer()

MySec = MySec + 1
If MySec = 60 Then
    MyMin = MyMin + 1
    MySec = 0
    If MyMin = 60 Then
        MyHour = MyHour + 1
        MyMin = 0
    End If
End If


If MySec < 10 Then
    lblTimeSec.Caption = "0" & MySec
Else: lblTimeSec.Caption = MySec
End If

If MyMin < 10 Then
    lblTimeMin.Caption = "0" & MyMin
Else: lblTimeMin.Caption = MyMin
End If

If MyHour < 10 Then
    lblTimeHour.Caption = "0" & MyHour
Else: lblTimeHour.Caption = MyHour
End If
    

End Sub

Private Sub tmrTruckCar_Timer()
    
    If CarWay(TruckCar) = "Start" Then     'if the car start a new line
        Call StartDriving(TruckCar, imgTruckCar, "Truck")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(TruckCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(TruckCar, imgTruckCar, TruckCarsquare, Truck, tmrTruckCar)
    End If
    
    If FirstInterval(TruckCar) = True Then  'Enable the next car interval
        tmrBlueCar.Enabled = True
        FirstInterval(TruckCar) = False
    End If

End Sub

Private Sub tmrWhiteCar_Timer()
    
    If CarWay(WhiteCar) = "Start" Then     'if the car start a new line
       Call StartDriving(WhiteCar, imgWhiteCar, "White")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(WhiteCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(WhiteCar, imgWhiteCar, WhiteCarsquare, Sport, tmrWhiteCar)
    End If
    
    If FirstInterval(WhiteCar) = True Then  'Enable the next car interval
        tmrYellowCar.Enabled = True
        FirstInterval(WhiteCar) = False
    End If
    
End Sub

Private Sub tmrYellowCar_Timer()
    
    If CarWay(YellowCar) = "Start" Then     'if the car start a new line
       Call StartDriving(YellowCar, imgYellowCar, "Yellow")
    End If
'-------------------------------------------------------------------------------------------------------
                ' START DRIVING
'-------------------------------------------------------------------------------------------------------
    If CarWay(YellowCar) = "Move" Then
        Call mdlMoveInLines.MoveInLine(YellowCar, imgYellowCar, YellowCarsquare, Sport, tmrYellowCar)
    End If
    
    If FirstInterval(YellowCar) = True Then  'Enable the next car interval
        tmrCallPolice.Enabled = True
        FirstInterval(YellowCar) = False
    End If
    
End Sub

Public Sub tmrYellowLight_Timer()

    If intMyLight < 8 Then
        intMyLight = intMyLight + 1
    Else: intMyLight = 0
    End If
    
    Select Case intMyLight
        Case 1
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 5
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 6
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 3
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 4
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 7
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor23 = "RED"
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 2
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor14 = "RED"
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
        Case 8
            imgLight14.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            imgLight23.Picture = LoadPicture(App.Path & "\Images\lightos.bmp")
            strLightColor4 = "RED"
            imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor5 = "RED"
            imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor8 = "RED"
            imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor9 = "RED"
            imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor15 = "RED"
            imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
            strLightColor22 = "RED"
            imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
    End Select
    
    boolYellow = False
    tmrYellowLight.Enabled = False
    tmrLights.Enabled = True
End Sub
