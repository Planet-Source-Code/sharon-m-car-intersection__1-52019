Attribute VB_Name = "mdlPolice"
Option Explicit
Option Base 1

Function PoliceLine() As Integer
    Dim PoliceInter As Integer
    Randomize
    
    PoliceInter = Int((12 * Rnd) + 1)
    If CarsInLine(PoliceInter) <= 0 Then
        PoliceLine = IntersectionNumber(PoliceInter)
    Else
        PoliceLine = PoliceLine()
    End If
    
End Function
Public Sub MakeLineBusy(LineSquare() As Boolean)
    Dim i As Integer
    For i = 1 To intHaveToSlow
            LineSquare(i) = True
        Next i
End Sub

Public Sub GoToStation()
    frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
    frmIntersection.imgPolice.Top = CarsStartPosition(9).Top
    frmIntersection.imgPolice.Left = CarsStartPosition(9).Left
    
End Sub

Public Sub SetPoliceCar()

     Select Case CarIntersection(Police)    'Setting the car with the right picture in the right coordinate
            Case 1
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(1, frmIntersection.imgPolice)
            Case 2
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(2, frmIntersection.imgPolice)
            Case 3
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(3, frmIntersection.imgPolice)
            Case 4
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(4, frmIntersection.imgPolice)
                CarsInLine(1) = CarsInLine(1) + 1
            Case 5
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(5, frmIntersection.imgPolice)
                CarsInLine(2) = CarsInLine(2) + 1
            Case 6
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(6, frmIntersection.imgPolice)
                CarsInLine(3) = CarsInLine(3) + 1
            Case 7
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(7, frmIntersection.imgPolice)
                CarsInLine(4) = CarsInLine(4) + 1
            Case 8
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(8, frmIntersection.imgPolice)
                CarsInLine(5) = CarsInLine(5) + 1
            Case 9
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_left.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(9, frmIntersection.imgPolice)
                CarsInLine(6) = CarsInLine(6) + 1
            Case 10
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(10, frmIntersection.imgPolice)
            Case 11
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(11, frmIntersection.imgPolice)
            Case 12
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_right.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(12, frmIntersection.imgPolice)
            Case 13
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(13, frmIntersection.imgPolice)
                CarsInLine(7) = CarsInLine(7) + 1
            Case 14
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(14, frmIntersection.imgPolice)
                CarsInLine(8) = CarsInLine(8) + 1
            Case 15
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(15, frmIntersection.imgPolice)
                CarsInLine(9) = CarsInLine(9) + 1
            Case 16
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(16, frmIntersection.imgPolice)
            Case 17
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(17, frmIntersection.imgPolice)
            Case 18
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(18, frmIntersection.imgPolice)
            Case 19
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(19, frmIntersection.imgPolice)
            Case 20
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(20, frmIntersection.imgPolice)
            Case 21
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_down.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(21, frmIntersection.imgPolice)
            Case 22
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(22, frmIntersection.imgPolice)
                CarsInLine(10) = CarsInLine(10) + 1
            Case 23
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(23, frmIntersection.imgPolice)
                CarsInLine(11) = CarsInLine(11) + 1
            Case 24
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_up.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(24, frmIntersection.imgPolice)
                CarsInLine(12) = CarsInLine(12) + 1
            Case 223
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_22.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(25, frmIntersection.imgPolice)
            Case 1510
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_15.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(26, frmIntersection.imgPolice)
            Case 416
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_4.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(28, frmIntersection.imgPolice)
            Case 921
                frmIntersection.imgPolice.Picture = LoadPicture(App.Path & "\Images\" & "Imergency" & "_9.bmp")
                CarWay(Police) = mdlCheckEntery.LinesCases(27, frmIntersection.imgPolice)
        End Select
        CarWay(Police) = "Move"
End Sub


Public Sub MovePolice(CarSquare As Integer)
    

Select Case CarIntersection(Police)

    Case 5
    'start line number 5 end at 11
        If ((CarIntersection(Police) = 5) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left <= Lines(5).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed    'Keep moving
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove      'Add k"m
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(5).LeftEnd Then      'When you at the end of the line
                    CarIntersection(Police) = 11       'Chang line number
                    CarWay(Police) = "Change"
                    Call mdlPolice.SetPoliceCar
                    LineBusy(2) = False
                    CarsInLine(2) = CarsInLine(2) - 1
                End If
            End If
        End If
        
    Case 11
        'start line number 11
        If ((CarIntersection(Police) = 11) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left <= Lines(11).LeftEnd) Then
                frmIntersection.imgPolice.Visible = True
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(11).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
    Case 8
        'start line number 8 end at 2
         If ((CarIntersection(Police) = 8) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left >= Lines(8).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(8).LeftEnd Then
                    CarIntersection(Police) = 2
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(5) = False
                    CarsInLine(5) = CarsInLine(5) - 1
                End If
            End If
        End If
        
     Case 2
        'start line number 2
        If ((CarIntersection(Police) = 2) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left >= Lines(2).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(2).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
     Case 14
        'start line number 14 end at 20
        If ((CarIntersection(Police) = 14) And (CarWay(Police) = "Move")) Then
             If (frmIntersection.imgPolice.Top <= Lines(14).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(14).TopEnd Then
                    CarIntersection(Police) = 20
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(8) = False
                    CarsInLine(8) = CarsInLine(8) - 1
                End If
            End If
        End If
        
    Case 20
        'start line number 20
        If ((CarIntersection(Police) = 20) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Top <= Lines(20).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(20).TopEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
    Case 23
        'start line number 23 end at 17
        If ((CarIntersection(Police) = 23) And (CarWay(Police) = "Move")) Then
             If (frmIntersection.imgPolice.Top >= Lines(23).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(23).TopEnd Then
                    CarIntersection(Police) = 17
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(11) = False
                    CarsInLine(11) = CarsInLine(11) - 1
                End If
            End If
        End If
        
    Case 17
        'start line number 17
        If ((CarIntersection(Police) = 17) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Top >= Lines(17).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(17).TopEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
     Case 6
        'start line number 6 move to 19
        If ((CarIntersection(Police) = 6) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left <= Lines(6).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(6).LeftEnd Then
                    CarIntersection(Police) = 19
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(3) = False
                    CarsInLine(3) = CarsInLine(3) - 1
                End If
            End If
        End If
        
    Case 19
        'start line number 19
        If ((CarIntersection(Police) = 19) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Top <= Lines(19).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(19).TopEnd Then
                    CarIntersection(Police) = 13
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
    Case 4
        'start line number 4 move to 16
        If ((CarIntersection(Police) = 4) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left <= Lines(4).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(4).LeftEnd Then
                    CarIntersection(Police) = 416
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(1) = False
                    CarsInLine(1) = CarsInLine(1) - 1
                End If
            End If
        End If
        
    Case 416
        'between 4 to 16
        If ((CarIntersection(Police) = 416) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left <= Lines(28).LeftEnd) Then
                If ((frmIntersection.imgPolice.Top <= Lines(28).TopStart) And (frmIntersection.imgPolice.Top >= Lines(28).TopEnd)) Then
                    frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - 30
                    frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + 32
                    MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                    If frmIntersection.imgPolice.Top - 30 <= Lines(28).TopEnd Then
                        CarIntersection(Police) = 16
                        CarWay(Police) = "Change"
                        Call SetPoliceCar
                    End If
                End If
            End If
        End If

Case 16
        'start line number 16
        If ((CarIntersection(Police) = 16) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Top >= Lines(16).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(16).TopEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
    
    Case 7
       'start line number 7 move to 18
        If ((CarIntersection(Police) = 7) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left >= Lines(7).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(7).LeftEnd Then
                    CarIntersection(Police) = 18
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(4) = False
                    CarsInLine(4) = CarsInLine(4) - 1
                End If
            End If
        End If
        
    Case 18
        'start line number 18
        If ((CarIntersection(Police) = 18) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Top <= Lines(18).TopStart) And (frmIntersection.imgPolice.Top >= Lines(18).TopEnd)) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(18).TopEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If

    Case 9
       'start line number 9 move to 21
        If ((CarIntersection(Police) = 9) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Left >= Lines(9).LeftEnd) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(9).LeftEnd Then
                    CarIntersection(Police) = 921
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(6) = False
                    CarsInLine(6) = CarsInLine(6) - 1
                End If
            End If
        End If
        
 Case 921
        'between 9 to 21
        If ((CarIntersection(Police) = 921) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left <= Lines(27).LeftStart) And (frmIntersection.imgPolice.Left >= Lines(27).LeftEnd)) Then
                If ((frmIntersection.imgPolice.Top >= Lines(27).TopStart) And (frmIntersection.imgPolice.Top <= Lines(27).TopEnd)) Then
                    frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + 24
                    frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - 24
                    MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                    If frmIntersection.imgPolice.Top + 24 >= Lines(27).TopEnd Then
                        CarIntersection(Police) = 21
                        CarWay(Police) = "Change"
                        Call SetPoliceCar
                    End If
                End If
            End If
        End If

Case 21
        'start line number 21
        If ((CarIntersection(Police) = 21) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Top >= Lines(21).TopStart) And (frmIntersection.imgPolice.Top <= Lines(21).TopEnd)) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(21).TopEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If

Case 13
        'start line number 13 move to 1
        If ((CarIntersection(Police) = 13) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Top >= Lines(13).TopStart) And (frmIntersection.imgPolice.Top <= Lines(13).TopEnd)) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(13).TopEnd Then
                    CarIntersection(Police) = 1
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(7) = False
                    CarsInLine(7) = CarsInLine(7) - 1
                End If
            End If
        End If

Case 1
        'start line number 1
        If ((CarIntersection(Police) = 1) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left <= Lines(1).LeftStart) And (frmIntersection.imgPolice.Left >= Lines(1).LeftEnd)) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(1).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
 Case 15
        'start line number 15 move to 10
        If ((CarIntersection(Police) = 15) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Top >= Lines(15).TopStart) And (frmIntersection.imgPolice.Top <= Lines(15).TopEnd)) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top + MyCars(Police).CarSpeed >= Lines(15).TopEnd Then
                    CarIntersection(Police) = 1510
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(9) = False
                    CarsInLine(9) = CarsInLine(9) - 1
                End If
            End If
        End If

Case 1510
        'between 15 to 10
        If ((CarIntersection(Police) = 1510) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left >= Lines(26).LeftStart) And (frmIntersection.imgPolice.Left <= Lines(26).LeftEnd)) Then
                If ((frmIntersection.imgPolice.Top >= Lines(26).TopStart) And (frmIntersection.imgPolice.Top <= Lines(26).TopEnd)) Then
                    frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top + 30
                    frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + 20
                    MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                    If frmIntersection.imgPolice.Top + 30 >= Lines(26).TopEnd Then
                        CarIntersection(Police) = 10
                        CarWay(Police) = "Change"
                        Call SetPoliceCar
                    End If
                End If
            End If
        End If

Case 10
        'start line number 10
        If ((CarIntersection(Police) = 10) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left >= Lines(10).LeftStart) And (frmIntersection.imgPolice.Left <= Lines(10).LeftEnd)) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(10).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If

Case 24
        'start line number 24 move to 12
        If ((CarIntersection(Police) = 24) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Top <= Lines(24).TopStart) And (frmIntersection.imgPolice.Top >= Lines(24).TopEnd)) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(24).TopEnd Then
                    CarIntersection(Police) = 12
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(12) = False
                    CarsInLine(12) = CarsInLine(12) - 1
                End If
            End If
        End If

Case 12
        'start line number 12
        If ((CarIntersection(Police) = 12) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left >= Lines(12).LeftStart) And (frmIntersection.imgPolice.Left <= Lines(12).LeftEnd)) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left + MyCars(Police).CarSpeed >= Lines(12).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
        
Case 22
        'start line number 22 move to 3
        If ((CarIntersection(Police) = 22) And (CarWay(Police) = "Move")) Then
            If (frmIntersection.imgPolice.Top >= Lines(22).TopEnd) Then
                frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Top - MyCars(Police).CarSpeed <= Lines(22).TopEnd Then
                    CarIntersection(Police) = 223
                    CarWay(Police) = "Change"
                    Call SetPoliceCar
                    LineBusy(10) = False
                    CarsInLine(10) = CarsInLine(10) - 1
                End If
            End If
        End If

Case 223
        'between 22 to 3
        If ((CarIntersection(Police) = 223) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left <= Lines(25).LeftStart) And (frmIntersection.imgPolice.Left >= Lines(25).LeftEnd)) Then
                If ((frmIntersection.imgPolice.Top <= Lines(25).TopStart) And (frmIntersection.imgPolice.Top >= Lines(25).TopEnd)) Then
                    frmIntersection.imgPolice.Top = frmIntersection.imgPolice.Top - 30
                    frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - 14
                    MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                    If frmIntersection.imgPolice.Top - 30 <= Lines(25).TopEnd Then
                        CarIntersection(Police) = 3
                        CarWay(Police) = "Change"
                        Call SetPoliceCar
                    End If
                End If
            End If
        End If

Case 3
        'start line number 3
        If ((CarIntersection(Police) = 3) And (CarWay(Police) = "Move")) Then
            If ((frmIntersection.imgPolice.Left <= Lines(3).LeftStart) And (frmIntersection.imgPolice.Left >= Lines(3).LeftEnd)) Then
                frmIntersection.imgPolice.Left = frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed
                MyCars(Police).KM = MyCars(Police).KM + Vehicle(Imergency).KmPerMove
                If frmIntersection.imgPolice.Left - MyCars(Police).CarSpeed <= Lines(3).LeftEnd Then
                    CarWay(Police) = "Start"
                    Call GoToStation
                    frmIntersection.tmrPolice.Enabled = False
                    frmIntersection.tmrLights.Enabled = True
                    Call TernLightGreen
                End If
            End If
        End If
    
    End Select
End Sub

Public Sub TernLightsRed()
    strLightColor8 = "RED"
    frmIntersection.imgLight8.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor9 = "RED"
    frmIntersection.imgLight9.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor4 = "RED"
    frmIntersection.imgLight4.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor5 = "RED"
    frmIntersection.imgLight5.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor14 = "RED"
    frmIntersection.imgLight14.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor23 = "RED"
    frmIntersection.imgLight23.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor15 = "RED"
    frmIntersection.imgLight15.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
    strLightColor22 = "RED"
    frmIntersection.imgLight22.Picture = LoadPicture(App.Path & "\Images\lightrs.bmp")
        
End Sub

Public Sub TernLightGreen()
    intMyLight = intMyLight - 1
    boolYellow = True
    frmIntersection.tmrLights.Enabled = True
    frmIntersection.tmrYellowLight.Enabled = True
End Sub
