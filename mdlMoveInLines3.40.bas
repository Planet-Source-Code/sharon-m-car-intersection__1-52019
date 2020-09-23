Attribute VB_Name = "mdlMoveInLines"
Option Explicit
Option Base 1

Public Sub MoveInLine(CarName As Integer, CarImage As Image, CarSquare As Integer, CarType As Integer, CarTimer As Timer)

Select Case CarIntersection(CarName)

Case 5
 'start line number 5 end at 11
        If ((CarIntersection(CarName) = 5) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(5).LeftEnd) Or (strLightColor5 = "GREEN")) Then
                 If LinesSquares5(CarSquare + intHaveToSlow) = True Then     'Checking if the way is free
                    CarImage.Left = CarImage.Left   'If not - stop
                Else    'If it's open
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed    'Keep moving
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer    'Lose fuel
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove      'Add k"m
                    CarSquare = CarSquare + 1       'Forword in the line array
                    LinesSquares5(CarSquare) = True     'Check the square you are in now
                    LinesSquares5(CarSquare - 1) = False    'Uncheck the last square
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(5).LeftEnd Then      'When you at the end of the line
                    If strLightColor5 = "GREEN" Then        'And the light is green
                        CarIntersection(CarName) = 11       'Chang line number
                        CarWay(CarName) = "Start"
                        LineBusy(2) = False
                        CarsInLine(2) = CarsInLine(2) - 1
                    Else
                        CarTimer.Enabled = False      'If the light is red - stop
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
 Case 11
        'start line number 11
        If ((CarIntersection(CarName) = 11) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Left <= Lines(11).LeftEnd) Then
                If LinesSquares5(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares5(CarSquare) = True
                    LinesSquares5(CarSquare - 1) = False
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(11).LeftEnd Then
                    CarSquare = MakeLineFalse(LinesSquares5(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 6
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12()
                        If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 8
        'start line number 8 end at 2
         If ((CarIntersection(CarName) = 8) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left >= Lines(8).LeftEnd) Or (strLightColor8 = "GREEN")) Then
                If LinesSquares8(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares8(CarSquare) = True
                    LinesSquares8(CarSquare - 1) = False
                End If
                If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(8).LeftEnd Then
                    If strLightColor8 = "GREEN" Then
                        CarIntersection(CarName) = 2
                        CarWay(CarName) = "Start"
                        LineBusy(5) = False
                        CarsInLine(5) = CarsInLine(5) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
 Case 2
        'start line number 2
        If ((CarIntersection(CarName) = 2) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Left >= Lines(2).LeftEnd) Then
                If LinesSquares8(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares8(CarSquare) = True
                    LinesSquares8(CarSquare - 1) = False
                End If
                If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(2).LeftEnd Then
                    CarSquare = MakeLineFalse(LinesSquares8(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 7
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                        If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
    
   Case 14
        'start line number 14 end at 20
        If ((CarIntersection(CarName) = 14) And (CarWay(CarName) = "Move")) Then
             If ((CarImage.Top <= Lines(14).TopEnd) Or (strLightColor14 = "GREEN")) Then
                If LinesSquares14(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares14(CarSquare) = True
                    LinesSquares14(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(14).TopEnd Then
                    If strLightColor14 = "GREEN" Then
                        CarIntersection(CarName) = 20
                        CarWay(CarName) = "Start"
                        LineBusy(8) = False
                        CarsInLine(8) = CarsInLine(8) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 20
        'start line number 20
        If ((CarIntersection(CarName) = 20) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Top <= Lines(20).TopEnd) Then
                If LinesSquares14(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares14(CarSquare) = True
                    LinesSquares14(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(20).TopEnd Then
                    CarSquare = MakeLineFalse(LinesSquares14(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 13
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        If CarsInLine(CarIntersection(CarName) - 6) >= 2 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
    
  Case 23
        'start line number 23 end at 17
        If ((CarIntersection(CarName) = 23) And (CarWay(CarName) = "Move")) Then
             If ((CarImage.Top >= Lines(23).TopEnd) Or (strLightColor23 = "GREEN")) Then
                If LinesSquares23(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares23(CarSquare) = True
                    LinesSquares23(CarSquare - 1) = False
                End If
                If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(23).TopEnd Then
                    If strLightColor23 = "GREEN" Then
                        CarIntersection(CarName) = 17
                        CarWay(CarName) = "Start"
                        LineBusy(11) = False
                        CarsInLine(11) = CarsInLine(11) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 17
        'start line number 17
        If ((CarIntersection(CarName) = 17) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Top >= Lines(17).TopEnd) Then
                If LinesSquares23(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares23(CarSquare) = True
                    LinesSquares23(CarSquare - 1) = False
                End If
                If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(17).TopEnd Then
                    CarSquare = MakeLineFalse(LinesSquares23(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 22
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                        If CarsInLine(CarIntersection(CarName) - 12) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
 Case 6
        'start line number 6 move to 19
        If ((CarIntersection(CarName) = 6) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Left <= Lines(6).LeftEnd) Then
               If LinesSquares6(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares6(CarSquare) = True
                    LinesSquares6(CarSquare - 1) = False
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(6).LeftEnd Then
                    CarIntersection(CarName) = 19
                    CarWay(CarName) = "Start"
                    LineBusy(3) = False
                    CarsInLine(3) = CarsInLine(3) - 1
                End If
            End If
        End If

Case 19
        'start line number 19
        If ((CarIntersection(CarName) = 19) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Top <= Lines(19).TopEnd) Then
                If LinesSquares6(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares6(CarSquare) = True
                    LinesSquares6(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(19).TopEnd Then
                    CarSquare = MakeLineFalse(LinesSquares6(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 13
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        If CarsInLine(CarIntersection(CarName) - 6) >= 2 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
 Case 4
        'start line number 4 move to 16
        If ((CarIntersection(CarName) = 4) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(4).LeftEnd) Or (strLightColor4 = "GREEN")) Then
               If LinesSquares4(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares4(CarSquare) = True
                    LinesSquares4(CarSquare - 1) = False
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(4).LeftEnd Then
                    If strLightColor4 = "GREEN" Then
                        CarIntersection(CarName) = 416
                        CarWay(CarName) = "Start"
                        LineBusy(1) = False
                        CarsInLine(1) = CarsInLine(1) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 416
        'between 4 to 16
        If ((CarIntersection(CarName) = 416) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Left <= Lines(28).LeftEnd) Then
                If ((CarImage.Top <= Lines(28).TopStart) And (CarImage.Top >= Lines(28).TopEnd)) Then
                     If LinesSquares4(CarSquare + intHaveToSlow) = True Then
                        CarImage.Top = CarImage.Top
                        CarImage.Left = CarImage.Left
                     Else
                        CarImage.Top = CarImage.Top - 30
                        CarImage.Left = CarImage.Left + 32
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares4(CarSquare) = True
                        LinesSquares4(CarSquare - 1) = False
                    End If
                    If CarImage.Top - 30 <= Lines(28).TopEnd Then
                        CarIntersection(CarName) = 16
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 16
        'start line number 16
        If ((CarIntersection(CarName) = 16) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Top >= Lines(16).TopEnd) Then
                If LinesSquares4(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares4(CarSquare) = True
                    LinesSquares4(CarSquare - 1) = False
                End If
                If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(16).TopEnd Then
                    CarSquare = MakeLineFalse(LinesSquares4(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 22
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                        If CarsInLine(CarIntersection(CarName) - 12) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 7
       'start line number 7 move to 18
        If ((CarIntersection(CarName) = 7) And (CarWay(CarName) = "Move")) Then
            If (CarImage.Left >= Lines(7).LeftEnd) Then
                If LinesSquares7(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares7(CarSquare) = True
                    LinesSquares7(CarSquare - 1) = False
                End If
                If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(7).LeftEnd Then
                    CarIntersection(CarName) = 18
                    CarWay(CarName) = "Start"
                    LineBusy(4) = False
                    CarsInLine(4) = CarsInLine(4) - 1
                End If
            End If
        End If
        
Case 18
        'start line number 18
        If ((CarIntersection(CarName) = 18) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top <= Lines(18).TopStart) And (CarImage.Top >= Lines(18).TopEnd)) Then
                If ((CarImage.Top - MyCars(CarName).CarSpeed <= FuelStop1.Top) And (CarNeedFuel(CarName) = True)) Then
                    If (boolFuel1Full = False) Then
                        boolFuel1Full = True
                        CarImage.Left = FuelStop1.Left
                        CarImage.Top = FuelStop1.Top
                        LinesSquares7(CarSquare) = False
                        CarIsFueling1 = CarName
                        frmIntersection.tmrFuel1.Enabled = True
                        mdlFuel.FuelLabel1 (CarName)
                    End If
                Else
                    If LinesSquares7(CarSquare + intHaveToSlow) = True Then
                        CarImage.Top = CarImage.Top
                    Else
                        CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares7(CarSquare) = True
                        LinesSquares7(CarSquare - 1) = False
                    End If
                    If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(18).TopEnd Then
                        CarSquare = MakeLineFalse(LinesSquares7(), CarSquare)
                        If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                            CarNeedFuel(CarName) = True
                            CarIntersection(CarName) = 22
                            CarWay(CarName) = "Start"
                        Else
                        'replacing line number
                            CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                            If CarsInLine(CarIntersection(CarName) - 12) >= 3 Then
                                CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18()
                            End If
                            CarWay(CarName) = "Start"
                        End If
                    End If
                End If
            End If
        End If

Case 9
       'start line number 9 move to 21
        If ((CarIntersection(CarName) = 9) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left >= Lines(9).LeftEnd) Or (strLightColor9 = "GREEN")) Then
                If LinesSquares9(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares9(CarSquare) = True
                    LinesSquares9(CarSquare - 1) = False
                End If
                If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(9).LeftEnd Then
                    If strLightColor9 = "GREEN" Then
                        CarIntersection(CarName) = 921
                        CarWay(CarName) = "Start"
                        LineBusy(6) = False
                        CarsInLine(6) = CarsInLine(6) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
 Case 921
        'between 9 to 21
        If ((CarIntersection(CarName) = 921) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(27).LeftStart) And (CarImage.Left >= Lines(27).LeftEnd)) Then
                If ((CarImage.Top >= Lines(27).TopStart) And (CarImage.Top <= Lines(27).TopEnd)) Then
                    If LinesSquares9(CarSquare + intHaveToSlow) = True Then
                        CarImage.Top = CarImage.Top
                        CarImage.Left = CarImage.Left
                    Else
                        CarImage.Top = CarImage.Top + 24
                        CarImage.Left = CarImage.Left - 24
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares9(CarSquare) = True
                        LinesSquares9(CarSquare - 1) = False
                    End If
                    If CarImage.Top + 24 >= Lines(27).TopEnd Then
                        CarIntersection(CarName) = 21
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 21
        'start line number 21
        If ((CarIntersection(CarName) = 21) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top >= Lines(21).TopStart) And (CarImage.Top <= Lines(21).TopEnd)) Then
               If LinesSquares9(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares9(CarSquare) = True
                    LinesSquares9(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(21).TopEnd Then
                    CarSquare = MakeLineFalse(LinesSquares9(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 13
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        If CarsInLine(CarIntersection(CarName) - 6) >= 2 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 13
        'start line number 13 move to 1
        If ((CarIntersection(CarName) = 13) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top >= Lines(13).TopStart) And (CarImage.Top <= Lines(13).TopEnd)) Then
                If LinesSquares13(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares13(CarSquare) = True
                    LinesSquares13(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(13).TopEnd Then
                    CarIntersection(CarName) = 1
                    CarWay(CarName) = "Start"
                    LineBusy(7) = False
                    CarsInLine(7) = CarsInLine(7) - 1
                End If
            End If
        End If

Case 1
        'start line number 1
        If ((CarIntersection(CarName) = 1) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(1).LeftStart) And (CarImage.Left >= Lines(1).LeftEnd)) Then
                If ((CarImage.Left - MyCars(CarName).CarSpeed <= FuelStop2.Left) And (CarNeedFuel(CarName) = True)) Then
                    If boolFuel2Full = False Then
                        boolFuel2Full = True
                        CarImage.Left = FuelStop2.Left
                        CarImage.Top = FuelStop2.Top
                        LinesSquares13(CarSquare) = False
                        CarIsFueling2 = CarName
                        frmIntersection.tmrFuel2.Enabled = True
                        mdlFuel.FuelLabel2 (CarName)
                    End If
                Else
                    If LinesSquares13(CarSquare + intHaveToSlow) = True Then
                        CarImage.Left = CarImage.Left
                    Else
                        CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares13(CarSquare) = True
                        LinesSquares13(CarSquare - 1) = False
                    End If
                    If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(1).LeftEnd Then
                        CarSquare = MakeLineFalse(LinesSquares13(), CarSquare)
                        If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                            CarNeedFuel(CarName) = True
                            CarIntersection(CarName) = 7
                            CarWay(CarName) = "Start"
                        Else
                        'replacing line number
                            CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                            If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                                CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                            End If
                            CarWay(CarName) = "Start"
                        End If
                    End If
                End If
            End If
        End If
        
 Case 15
        'start line number 15 move to 10
        If ((CarIntersection(CarName) = 15) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top <= Lines(15).TopEnd) Or (strLightColor15 = "GREEN")) Then
                If LinesSquares15(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares15(CarSquare) = True
                    LinesSquares15(CarSquare - 1) = False
                End If
                If CarImage.Top + MyCars(CarName).CarSpeed >= Lines(15).TopEnd Then
                    If strLightColor15 = "GREEN" Then
                        CarIntersection(CarName) = 1510
                        CarWay(CarName) = "Start"
                        LineBusy(9) = False
                        CarsInLine(9) = CarsInLine(9) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 1510
        'between 15 to 10
        If ((CarIntersection(CarName) = 1510) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left >= Lines(26).LeftStart) And (CarImage.Left <= Lines(26).LeftEnd)) Then
                If ((CarImage.Top >= Lines(26).TopStart) And (CarImage.Top <= Lines(26).TopEnd)) Then
                    If LinesSquares15(CarSquare + intHaveToSlow) = True Then
                        CarImage.Top = CarImage.Top
                        CarImage.Left = CarImage.Left
                    Else
                        CarImage.Top = CarImage.Top + 30
                        CarImage.Left = CarImage.Left + 20
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares15(CarSquare) = True
                        LinesSquares15(CarSquare - 1) = False
                    End If
                    If CarImage.Top + 30 >= Lines(26).TopEnd Then
                        CarIntersection(CarName) = 10
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 10
        'start line number 10
        If ((CarIntersection(CarName) = 10) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left >= Lines(10).LeftStart) And (CarImage.Left <= Lines(10).LeftEnd)) Then
                If LinesSquares15(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares15(CarSquare) = True
                    LinesSquares15(CarSquare - 1) = False
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(10).LeftEnd Then
                    CarSquare = MakeLineFalse(LinesSquares15(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 6
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12
                        If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 24
        'start line number 24 move to 12
        If ((CarIntersection(CarName) = 24) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top <= Lines(24).TopStart) And (CarImage.Top >= Lines(24).TopEnd)) Then
                If LinesSquares24(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares24(CarSquare) = True
                    LinesSquares24(CarSquare - 1) = False
                End If
                If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(24).TopEnd Then
                    CarIntersection(CarName) = 12
                    CarWay(CarName) = "Start"
                    LineBusy(12) = False
                    CarsInLine(12) = CarsInLine(12) - 1
                End If
            End If
        End If

Case 12
        'start line number 12
        If ((CarIntersection(CarName) = 12) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left >= Lines(12).LeftStart) And (CarImage.Left <= Lines(12).LeftEnd)) Then
                 If LinesSquares24(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left + MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares24(CarSquare) = True
                    LinesSquares24(CarSquare - 1) = False
                End If
                If CarImage.Left + MyCars(CarName).CarSpeed >= Lines(12).LeftEnd Then
                    CarSquare = MakeLineFalse(LinesSquares24(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 6
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12
                        If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
Case 22
        'start line number 22 move to 3
        If ((CarIntersection(CarName) = 22) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Top >= Lines(22).TopEnd) Or (strLightColor22 = "GREEN")) Then
                If LinesSquares22(CarSquare + intHaveToSlow) = True Then
                    CarImage.Top = CarImage.Top
                Else
                    CarImage.Top = CarImage.Top - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares22(CarSquare) = True
                    LinesSquares22(CarSquare - 1) = False
                End If
                If CarImage.Top - MyCars(CarName).CarSpeed <= Lines(22).TopEnd Then
                    If strLightColor22 = "GREEN" Then
                        CarIntersection(CarName) = 223
                        CarWay(CarName) = "Start"
                        LineBusy(10) = False
                        CarsInLine(10) = CarsInLine(10) - 1
                    Else
                        CarTimer.Enabled = False
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 223
        'between 22 to 3
        If ((CarIntersection(CarName) = 223) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(25).LeftStart) And (CarImage.Left >= Lines(25).LeftEnd)) Then
                If ((CarImage.Top <= Lines(25).TopStart) And (CarImage.Top >= Lines(25).TopEnd)) Then
                    If LinesSquares22(CarSquare + intHaveToSlow) = True Then
                        CarImage.Top = CarImage.Top
                        CarImage.Left = CarImage.Left
                    Else
                        CarImage.Top = CarImage.Top - 30
                        CarImage.Left = CarImage.Left - 14
                        MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                        MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                        CarSquare = CarSquare + 1
                        LinesSquares22(CarSquare) = True
                        LinesSquares22(CarSquare - 1) = False
                    End If
                    If CarImage.Top - 30 <= Lines(25).TopEnd Then
                        CarIntersection(CarName) = 3
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If

Case 3
        'start line number 3
        If ((CarIntersection(CarName) = 3) And (CarWay(CarName) = "Move")) Then
            If ((CarImage.Left <= Lines(3).LeftStart) And (CarImage.Left >= Lines(3).LeftEnd)) Then
                If LinesSquares22(CarSquare + intHaveToSlow) = True Then
                    CarImage.Left = CarImage.Left
                Else
                    CarImage.Left = CarImage.Left - MyCars(CarName).CarSpeed
                    MyCars(CarName).Fuel.FuelForNow = MyCars(CarName).Fuel.FuelForNow - FuelPerTimer
                    MyCars(CarName).KM = MyCars(CarName).KM + Vehicle(CarType).KmPerMove
                    CarSquare = CarSquare + 1
                    LinesSquares22(CarSquare) = True
                    LinesSquares22(CarSquare - 1) = False
                End If
                If CarImage.Left - MyCars(CarName).CarSpeed <= Lines(3).LeftEnd Then
                    CarSquare = MakeLineFalse(LinesSquares22(), CarSquare)
                    If MyCars(CarName).Fuel.FuelForNow < intNeedFuel Then
                        CarNeedFuel(CarName) = True
                        CarIntersection(CarName) = 7
                        CarWay(CarName) = "Start"
                    Else
                    'replacing line number
                        CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                        If CarsInLine(CarIntersection(CarName) - 3) >= 3 Then
                            CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3()
                        End If
                        CarWay(CarName) = "Start"
                    End If
                End If
            End If
        End If
        
    

End Select

End Sub
