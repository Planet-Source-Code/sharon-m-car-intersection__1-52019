Attribute VB_Name = "mdlStartInit"
Option Explicit
Option Base 1
Public Lines(28) As Lines_Type
Public IntersectionNumber(12) As Integer
Public Vehicle(4) As Vehicle_Type
Public MyCars(9) As Cars_Type
Public CarNeedFuel(8) As Boolean

Public Sub NewGameInitialazetion()  'Set all the lines sqaure to be false
    Dim i As Integer

    For i = 1 To SquaresNum
        LinesSquares4(i) = False
        LinesSquares5(i) = False
        LinesSquares6(i) = False
        LinesSquares7(i) = False
        LinesSquares8(i) = False
        LinesSquares9(i) = False
        LinesSquares13(i) = False
        LinesSquares14(i) = False
        LinesSquares15(i) = False
        LinesSquares22(i) = False
        LinesSquares23(i) = False
        LinesSquares24(i) = False
        LinesSquares25(i) = False
    Next i
    
    For i = 1 To 12
        LineBusy(i) = False
        CarsInLine(i) = 0
    Next i
    
'----------------------------------------------------------------
'   First Intersection choosing
'----------------------------------------------------------------

'Blue Car
        MyIntersection(BlueCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(BlueCar) = IntersectionNumber(MyIntersection(BlueCar))
        CarWay(BlueCar) = "Start"
        
'Fire Car
        MyIntersection(FireCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(FireCar) = IntersectionNumber(MyIntersection(FireCar))
        CarWay(FireCar) = "Start"

'Yellow Car
        MyIntersection(YellowCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(YellowCar) = IntersectionNumber(MyIntersection(YellowCar))
        CarWay(YellowCar) = "Start"
        
'Red Car
        MyIntersection(RedCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(RedCar) = IntersectionNumber(MyIntersection(RedCar))
        CarWay(RedCar) = "Start"
        
'White Car
        MyIntersection(WhiteCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(WhiteCar) = IntersectionNumber(MyIntersection(WhiteCar))
        CarWay(WhiteCar) = "Start"

'Black Car
        MyIntersection(BlackCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(BlackCar) = IntersectionNumber(MyIntersection(BlackCar))
        CarWay(BlackCar) = "Start"
        
'Green Car
        MyIntersection(GreenCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(GreenCar) = IntersectionNumber(MyIntersection(GreenCar))
        CarWay(GreenCar) = "Start"
        
'Truck Car
        MyIntersection(TruckCar) = Int((12 * Rnd) + 1)   ' Generate random value between 1 and 12.
        CarIntersection(TruckCar) = IntersectionNumber(MyIntersection(TruckCar))
        CarWay(TruckCar) = "Start"

'Police
        CarWay(Police) = "Start"
'----------------------------------------------------------------
    'Cars initializing
'----------------------------------------------------------------

'Blue Car
    MyCars(BlueCar).WhatCar = Vehicle(Normal).Type
    MyCars(BlueCar).CarSpeed = Vehicle(Normal).Speed
    MyCars(BlueCar).Fuel.FullTank = Vehicle(Normal).FullTank
    MyCars(BlueCar).Fuel.FuelForNow = Int((15000 * Rnd) + 5000)
    MyCars(BlueCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(BlueCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(BlueCar) = True
    Else: CarNeedFuel(BlueCar) = False
    End If
    
'Fire Car
    MyCars(FireCar).WhatCar = Vehicle(Sport).Type
    MyCars(FireCar).CarSpeed = Vehicle(Sport).Speed
    MyCars(FireCar).Fuel.FullTank = Vehicle(Sport).FullTank
    MyCars(FireCar).Fuel.FuelForNow = Int((20000 * Rnd) + 5000)
    MyCars(FireCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(FireCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(FireCar) = True
    Else: CarNeedFuel(FireCar) = False
    End If
    
'Yellow Car
    MyCars(YellowCar).WhatCar = Vehicle(Sport).Type
    MyCars(YellowCar).CarSpeed = Vehicle(Sport).Speed
    MyCars(YellowCar).Fuel.FullTank = Vehicle(Sport).FullTank
    MyCars(YellowCar).Fuel.FuelForNow = Int((20000 * Rnd) + 5000)
    MyCars(YellowCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(YellowCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(YellowCar) = True
    Else: CarNeedFuel(YellowCar) = False
    End If
    
'Truck Car
    MyCars(TruckCar).WhatCar = Vehicle(Truck).Type
    MyCars(TruckCar).CarSpeed = Vehicle(Truck).Speed
    MyCars(TruckCar).Fuel.FullTank = Vehicle(Truck).FullTank
    MyCars(TruckCar).Fuel.FuelForNow = Int((25000 * Rnd) + 5000)
    MyCars(TruckCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(TruckCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(TruckCar) = True
    Else: CarNeedFuel(TruckCar) = False
    End If
    
'Red Car
    MyCars(RedCar).WhatCar = Vehicle(Normal).Type
    MyCars(RedCar).CarSpeed = Vehicle(Normal).Speed
    MyCars(RedCar).Fuel.FullTank = Vehicle(Normal).FullTank
    MyCars(RedCar).Fuel.FuelForNow = Int((15000 * Rnd) + 5000)
    MyCars(RedCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(RedCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(RedCar) = True
    Else: CarNeedFuel(RedCar) = False
    End If

'White Car
    MyCars(WhiteCar).WhatCar = Vehicle(Sport).Type
    MyCars(WhiteCar).CarSpeed = Vehicle(Sport).Speed
    MyCars(WhiteCar).Fuel.FullTank = Vehicle(Sport).FullTank
    MyCars(WhiteCar).Fuel.FuelForNow = Int((20000 * Rnd) + 5000)
    MyCars(WhiteCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(WhiteCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(WhiteCar) = True
    Else: CarNeedFuel(WhiteCar) = False
    End If
    
'Black Car
    MyCars(BlackCar).WhatCar = Vehicle(Normal).Type
    MyCars(BlackCar).CarSpeed = Vehicle(Normal).Speed
    MyCars(BlackCar).Fuel.FullTank = Vehicle(Normal).FullTank
    MyCars(BlackCar).Fuel.FuelForNow = Int((15000 * Rnd) + 5000)
    MyCars(BlackCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(BlackCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(BlackCar) = True
    Else: CarNeedFuel(BlackCar) = False
    End If

'Green Car
    MyCars(GreenCar).WhatCar = Vehicle(Normal).Type
    MyCars(GreenCar).CarSpeed = Vehicle(Normal).Speed
    MyCars(GreenCar).Fuel.FullTank = Vehicle(Normal).FullTank
    MyCars(GreenCar).Fuel.FuelForNow = Int((15000 * Rnd) + 5000)
    MyCars(GreenCar).KM = Int((3000 * Rnd) + 1)
    
    If MyCars(GreenCar).Fuel.FuelForNow < intNeedFuel Then
        CarNeedFuel(GreenCar) = True
    Else: CarNeedFuel(GreenCar) = False
    End If

'Police Car
    MyCars(Police).WhatCar = Vehicle(Imergency).Type
    MyCars(Police).CarSpeed = Vehicle(Imergency).Speed
    MyCars(Police).Fuel.FullTank = Vehicle(Imergency).FullTank
    MyCars(Police).Fuel.FuelForNow = Int((17000 * Rnd) + 3000)
    MyCars(Police).KM = Int((3000 * Rnd) + 1)
        
End Sub


Public Sub FormInitialazetion()
    
'------------------------------------------------------------------
    'line numbers - for rundom choos
'------------------------------------------------------------------

    IntersectionNumber(1) = 4
    IntersectionNumber(2) = 5
    IntersectionNumber(3) = 6
    IntersectionNumber(4) = 7
    IntersectionNumber(5) = 8
    IntersectionNumber(6) = 9
    IntersectionNumber(7) = 13
    IntersectionNumber(8) = 14
    IntersectionNumber(9) = 15
    IntersectionNumber(10) = 22
    IntersectionNumber(11) = 23
    IntersectionNumber(12) = 24
    
'------------------------------------------------------------------
    
    'lines initialization
        
'------------------------------------------------------------------
    
    Lines(1).TopStart = 2880
    Lines(1).TopEnd = 2880
    Lines(1).LeftStart = 2160
    Lines(1).LeftEnd = -500
    
    Lines(2).TopStart = 3360
    Lines(2).TopEnd = 3360
    Lines(2).LeftStart = 6240
    Lines(2).LeftEnd = -500
    
    Lines(3).TopStart = 4080
    Lines(3).TopEnd = 4080
    Lines(3).LeftStart = 2520
    Lines(3).LeftEnd = -500
    
    Lines(4).TopStart = 4700
    Lines(4).TopEnd = 4700
    Lines(4).LeftStart = -500
    Lines(4).LeftEnd = 2040
    
    Lines(5).TopStart = 5280
    Lines(5).TopEnd = 5280
    Lines(5).LeftStart = -500
    Lines(5).LeftEnd = 2040
    
    Lines(6).TopStart = 6000
    Lines(6).TopEnd = 6000
    Lines(6).LeftStart = -500
    Lines(6).LeftEnd = 2040
    
    Lines(7).TopStart = 2750
    Lines(7).TopEnd = 2750
    Lines(7).LeftStart = 9120
    Lines(7).LeftEnd = 5880
    
    Lines(8).TopStart = 3360
    Lines(8).TopEnd = 3360
    Lines(8).LeftStart = 9120
    Lines(8).LeftEnd = 6300
    
    Lines(9).TopStart = 4080
    Lines(9).TopEnd = 4080
    Lines(9).LeftStart = 9120
    Lines(9).LeftEnd = 6240
    
    Lines(10).TopStart = 4700
    Lines(10).TopEnd = 4700
    Lines(10).LeftStart = 5880
    Lines(10).LeftEnd = 9360
    
    Lines(11).TopStart = 5280
    Lines(11).TopEnd = 5280
    Lines(11).LeftStart = 2040
    Lines(11).LeftEnd = 9360
    
    Lines(12).TopStart = 6000
    Lines(12).TopEnd = 6000
    Lines(12).LeftStart = 6000
    Lines(12).LeftEnd = 9360
    
    Lines(13).TopStart = -360
    Lines(13).TopEnd = 2040
    Lines(13).LeftStart = 2800
    Lines(13).LeftEnd = 2800
    
    Lines(14).TopStart = -360
    Lines(14).TopEnd = 1800
    Lines(14).LeftStart = 3400
    Lines(14).LeftEnd = 3400
    
    Lines(15).TopStart = -360
    Lines(15).TopEnd = 1800
    Lines(15).LeftStart = 3960
    Lines(15).LeftEnd = 3960
    
    Lines(16).TopStart = 2400
    Lines(16).TopEnd = -360
    Lines(16).LeftStart = 4600
    Lines(16).LeftEnd = 4600
    
    Lines(17).TopStart = 6400
    Lines(17).TopEnd = -360
    Lines(17).LeftStart = 5160
    Lines(17).LeftEnd = 5160
    
    Lines(18).TopStart = 2160
    Lines(18).TopEnd = -360
    Lines(18).LeftStart = 5760
    Lines(18).LeftEnd = 5760
    
    Lines(19).TopStart = 6240
    Lines(19).TopEnd = 9000
    Lines(19).LeftStart = 2800
    Lines(19).LeftEnd = 2800
       
    Lines(20).TopStart = 1800
    Lines(20).TopEnd = 9000
    Lines(20).LeftStart = 3400
    Lines(20).LeftEnd = 3400
    
    Lines(21).TopStart = 6000
    Lines(21).TopEnd = 9000
    Lines(21).LeftStart = 3960
    Lines(21).LeftEnd = 3960
    
    Lines(22).TopStart = 9000
    Lines(22).TopEnd = 6400
    Lines(22).LeftStart = 4600
    Lines(22).LeftEnd = 4600
    
    Lines(23).TopStart = 9000
    Lines(23).TopEnd = 6400
    Lines(23).LeftStart = 5160
    Lines(23).LeftEnd = 5160
    
    Lines(24).TopStart = 9000
    Lines(24).TopEnd = 6400
    Lines(24).LeftStart = 5760
    Lines(24).LeftEnd = 5760
    
    'From line 22 to 3
    Lines(25).TopStart = 6120
    Lines(25).TopEnd = 4320
    Lines(25).LeftStart = 4440
    Lines(25).LeftEnd = 3600
     
    'From line 15 to 10
    Lines(26).TopStart = 2760
    Lines(26).TopEnd = 4320
    Lines(26).LeftStart = 4420
    Lines(26).LeftEnd = 5520
     
    'From line 9 to 21
    Lines(27).TopStart = 4320
    Lines(27).TopEnd = 5760
    Lines(27).LeftStart = 5760
    Lines(27).LeftEnd = 4320
    
    'From line 4 to 16
    Lines(28).TopStart = 4320
    Lines(28).TopEnd = 3000
    Lines(28).LeftStart = 2760
    Lines(28).LeftEnd = 4200
    
    FuelStop1.Top = 400
    FuelStop1.Left = 7400
    
    FuelStop2.Top = 2040
    FuelStop2.Left = 960
    
    '==========================================================
    'Cars start position
    '============================================================
    
    CarsStartPosition(BlueCar).Top = 6960
    CarsStartPosition(BlueCar).Left = 0
    CarsStartPosition(BlueCar).Image = "\Images\blue_right.bmp"
    
    CarsStartPosition(FireCar).Top = 7560
    CarsStartPosition(FireCar).Left = 0
    CarsStartPosition(FireCar).Image = "\Images\fire_right.bmp"
    
    CarsStartPosition(RedCar).Top = 8040
    CarsStartPosition(RedCar).Left = 0
    CarsStartPosition(RedCar).Image = "\Images\red_right.bmp"
    
    CarsStartPosition(WhiteCar).Top = 8520
    CarsStartPosition(WhiteCar).Left = 0
    CarsStartPosition(WhiteCar).Image = "\Images\white_right.bmp"
    
    CarsStartPosition(BlackCar).Top = 6960
    CarsStartPosition(BlackCar).Left = 840
    CarsStartPosition(BlackCar).Image = "\Images\black_right.bmp"
    
    CarsStartPosition(GreenCar).Top = 9000
    CarsStartPosition(GreenCar).Left = 0
    CarsStartPosition(GreenCar).Image = "\Images\green_right.bmp"
    
    CarsStartPosition(YellowCar).Top = 7560
    CarsStartPosition(YellowCar).Left = 1440
    CarsStartPosition(YellowCar).Image = "\Images\yellow_down.bmp"
    
    CarsStartPosition(TruckCar).Top = 7440
    CarsStartPosition(TruckCar).Left = 960
    CarsStartPosition(TruckCar).Image = "\Images\truck_down.bmp"
    
    CarsStartPosition(Police).Top = 7440
    CarsStartPosition(Police).Left = 7560
    CarsStartPosition(Police).Image = "\Images\imergency_left.bmp"
    
    

'--------------------------------------------------------------
    'Vehicles settings
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
    Vehicle(3).Speed = 45
    Vehicle(3).FullTank = 30000
    Vehicle(3).KmPerMove = 70
   
    Vehicle(4).Type = "Imergency Car"
    Vehicle(4).Speed = 70
    Vehicle(4).FullTank = 20000
    Vehicle(4).KmPerMove = 100

    
End Sub

 Function ChangeLines19_20_21() As Integer  ' Generate random value between 13 to 14 - for changeing lines.

    Dim intRndNumber As Integer
    Dim CarIntersection As Integer
    
    intRndNumber = Int((3 * Rnd) + 1)
    Select Case intRndNumber
        Case 1
            CarIntersection = 13
        Case 2
            CarIntersection = 14
        Case 3
            CarIntersection = 15
    End Select
     
    ChangeLines19_20_21 = CarIntersection
    
End Function

Function ChangeLines10_11_12() As Integer     ' Generate random value between 4 to 6 - for changeing lines.

    Dim intRndNumber As Integer
    Dim CarIntersection As Integer
    Dim i As Integer
    
    intRndNumber = Int((3 * Rnd) + 1)
    Select Case intRndNumber
        Case 1
            CarIntersection = 4
        Case 2
            CarIntersection = 5
        Case 3
            CarIntersection = 6
    End Select
     
    ChangeLines10_11_12 = CarIntersection
    
End Function

Function ChangeLines1_2_3() As Integer   ' Generate random value between 7 to 9 - for changeing lines.

    Dim intRndNumber As Integer
    Dim CarIntersection As Integer
    
    intRndNumber = Int((3 * Rnd) + 1)
    Select Case intRndNumber
        Case 1
            CarIntersection = 7
        Case 2
            CarIntersection = 8
        Case 3
            CarIntersection = 9
    End Select
     
    ChangeLines1_2_3 = CarIntersection
    
End Function

Function ChangeLines16_17_18() As Integer    ' Generate random value between 22 to 24 - for changeing lines.

    Dim intRndNumber As Integer
    Dim CarIntersection As Integer
    
    intRndNumber = Int((3 * Rnd) + 1)
    Select Case intRndNumber
        Case 1
            CarIntersection = 22
        Case 2
            CarIntersection = 23
        Case 3
            CarIntersection = 24
    End Select
     
     
    ChangeLines16_17_18 = CarIntersection
    
End Function



'This function making the lines sqaure to be false after the car chang to other line
Function MakeLineFalse(LineSquare() As Boolean, CarSquare As Integer) As Integer
    Dim i As Integer
    
    For i = (CarSquare - 1) To SquaresNum
        LineSquare(i) = False
    Next i
    CarSquare = 1
    
    MakeLineFalse = CarSquare
End Function


