Attribute VB_Name = "mdlSettings"
Option Explicit
Option Base 1
Public Const SquaresNum = 350
Public Const BlueCar = 1
Public Const FireCar = 2
Public Const RedCar = 3
Public Const WhiteCar = 4
Public Const BlackCar = 5
Public Const GreenCar = 6
Public Const YellowCar = 7
Public Const TruckCar = 8
Public Const Police = 9
Public Const Sport = 1
Public Const Normal = 2
Public Const Truck = 3
Public Const Imergency = 4
Public fLen As Integer
Public filepath As String


    Type Lines_Type     'Definition of the lines
        TopStart As Integer
        TopEnd As Integer
        LeftStart As Integer
        LeftEnd As Integer
    End Type
    
    Type Vehicle_Type       'Definition of vehicle type
        Type As String
        FullTank As Integer
        Speed As Integer
        KmPerMove As Integer
    End Type
    
    
    Type Fuel_Type      'Definition of fuel type
        FullTank As Integer
        FuelForNow As Integer
    End Type
    
    
    Type Cars_Type      'Definition of each car
        WhatCar As String
        CarName As String
        CarSpeed As Integer
        Fuel As Fuel_Type
        KM As Long
    End Type
    
    Type StartPosition
        Top As Integer
        Left As Integer
        Image As String
    End Type
     

    Public CarIntersection(9) As Integer   'In what intersection is the car now
    Public CarWay(9) As String   'In what situation  car is in
    Public MyIntersection(8) As Integer
    Public DelayCounter(8) As Integer
    Public CarsStartPosition(9) As StartPosition
    
    Public LineBusy(12) As Boolean  'True if there is car in this line, False - if no car is in this line
    Public CarsInLine(12) As Integer
    
    Public intMyLight As Integer        'What line number should work
    
    Public strLightColor4 As String     'colors of the lights
    Public strLightColor5 As String
    Public strLightColor8 As String
    Public strLightColor9 As String
    Public strLightColor14 As String
    Public strLightColor15 As String
    Public strLightColor22 As String
    Public strLightColor23 As String
    
    
    Type FuelPoint      'Definition of fuel stop points
        Top As Integer
        Left As Integer
    End Type
    
    Public intNeedFuel As Integer       'When the cars should go to fuel
    Public FuelPerTimer As Integer
    
    Public FuelStop1 As FuelPoint       'Left and top points of the gas station
    Public FuelStop2 As FuelPoint
    
    Public boolFuel1Full As Boolean     'Is fuel stop one is busy
    Public boolFuel2Full As Boolean     'Is fuel stop two is busy
    
    Public CarIsFueling1 As Integer
    Public CarIsFueling2 As Integer
    
    Public MySec As Integer
    Public MyMin As Integer
    Public MyHour As Integer

    
    Type CarSavingDetails
        CarNumber As Integer
        CarFuel As Integer
        CarKM As Long
        CarTop As Integer
        CarLeft As Integer
        CarIntersection As Integer
        CarWay As String * 5
        CarSquare As Integer
    End Type
    
    Type Saving_Type        'The details we need for saving
        CarsDetails(9) As CarSavingDetails
        SaveLineBusy(12) As Boolean
        SaveCarsInLine(12) As Integer
        Fuel1IsWorking As Boolean
        Fuel2IsWorking As Boolean
        linesSquares(12, SquaresNum) As Boolean
    End Type
    
    Public SavingDetails As Saving_Type
        
    Public CarNumber As Integer     'How many cars are playing
    Public boolYellow As Boolean    'Light should chang to yellow
    Public intLightYellow As Integer        'Which traffic light should chang to yellow
    Public intHaveToSlow As Integer     'The space between two cars that will make them slow
    Public intAdd2Car As Integer        'Making the space between two cars bigger
    Public boolSaveInRest(9) As Boolean    'Check if the saving was doing while not driving
    Public intCarsInRest As Integer
    

    Public LinesSquares4(SquaresNum) As Boolean
    Public LinesSquares5(SquaresNum) As Boolean
    Public LinesSquares6(SquaresNum) As Boolean
    Public LinesSquares7(SquaresNum) As Boolean
    Public LinesSquares8(SquaresNum) As Boolean
    Public LinesSquares9(SquaresNum) As Boolean
    Public LinesSquares13(SquaresNum) As Boolean
    Public LinesSquares14(SquaresNum) As Boolean
    Public LinesSquares15(SquaresNum) As Boolean
    Public LinesSquares22(SquaresNum) As Boolean
    Public LinesSquares23(SquaresNum) As Boolean
    Public LinesSquares24(SquaresNum) As Boolean
    Public LinesSquares25(SquaresNum) As Boolean

    
