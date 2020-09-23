Attribute VB_Name = "mdlSetCases"
Option Explicit
Option Base 1

    Public Sub StartDriving(CarName As Integer, CarImage As Image, CarPic As String) ' As Boolean
        
        Select Case CarIntersection(CarName)    'Setting the car with the right picture in the right coordinate
            Case 1
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(1, CarImage)
            Case 2
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(2, CarImage)
            Case 3
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(3, CarImage)
            Case 4
                If CarsInLine(1) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 1, 4, LinesSquares4()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                        CarsInLine(1) = CarsInLine(1) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12
                End If
            Case 5
                If CarsInLine(2) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 2, 5, LinesSquares5()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                        CarsInLine(2) = CarsInLine(2) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12
                End If
            Case 6
                If CarsInLine(3) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 3, 6, LinesSquares6()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                        CarsInLine(3) = CarsInLine(3) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines10_11_12
                End If
            Case 7
                If CarsInLine(4) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 4, 7, LinesSquares7()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                        CarsInLine(4) = CarsInLine(4) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3
                End If
            Case 8
                If CarsInLine(5) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 5, 8, LinesSquares8()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                        CarsInLine(5) = CarsInLine(5) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3
                End If
            Case 9
                If CarsInLine(6) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 6, 9, LinesSquares9()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
                        CarsInLine(6) = CarsInLine(6) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines1_2_3
                End If
            Case 10
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(10, CarImage)
            Case 11
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(11, CarImage)
            Case 12
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(12, CarImage)
            Case 13
                If CarsInLine(7) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 7, 13, LinesSquares13()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                        CarsInLine(7) = CarsInLine(7) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21
                End If
            Case 14
                If CarsInLine(8) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 8, 14, LinesSquares14()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                        CarsInLine(8) = CarsInLine(8) + 1
                    End If
                Else
                     CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21
                End If
            Case 15
                If CarsInLine(9) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 9, 15, LinesSquares15()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                        CarsInLine(9) = CarsInLine(9) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines19_20_21
                End If
            Case 16
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(16, CarImage)
            Case 17
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(17, CarImage)
            Case 18
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(18, CarImage)
            Case 19
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(19, CarImage)
            Case 20
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(20, CarImage)
            Case 21
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(21, CarImage)
            Case 22
                If CarsInLine(10) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 10, 22, LinesSquares22()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                        CarsInLine(10) = CarsInLine(10) + 1
                    End If
                Else
                     CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18
                End If
            Case 23
                If CarsInLine(11) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 11, 23, LinesSquares23()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                        CarsInLine(11) = CarsInLine(11) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18
                End If
            Case 24
                If CarsInLine(12) <= 2 Then
                    If mdlCheckEntery.StartLinesCases(CarName, CarImage, 12, 24, LinesSquares24()) = True Then
                        CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
                        CarsInLine(12) = CarsInLine(12) + 1
                    End If
                Else
                    CarIntersection(CarName) = mdlStartInit.ChangeLines16_17_18
                End If
            Case 223
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_22.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(25, CarImage)
            Case 1510
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_15.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(26, CarImage)
            Case 416
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_4.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(28, CarImage)
            Case 921
                CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_9.bmp")
                CarWay(CarName) = mdlCheckEntery.LinesCases(27, CarImage)
        End Select

End Sub
