Attribute VB_Name = "mdlSaveNLoad"
Option Explicit
Option Base 1

Public Sub fnLoadingDetails(CarName As Integer, CarImage As Image, CarPic As String, CarTimer As Timer)          'Loading the car details of the last saved game

    If ((CarImage.Left = CarsStartPosition(CarName).Left) And (CarImage.Top = CarsStartPosition(CarName).Top) And (intCarsInRest = 9)) Then
        CarWay(CarName) = "End"
        CarImage.Left = CarsStartPosition(CarName).Left
        CarImage.Top = CarsStartPosition(CarName).Top
        CarImage.Picture = LoadPicture(App.Path & CarsStartPosition(CarName).Image)
        CarTimer.Enabled = False
        
    Else
        MyCars(CarName).Fuel.FuelForNow = SavingDetails.CarsDetails(CarName).CarFuel
        CarIntersection(CarName) = SavingDetails.CarsDetails(CarName).CarIntersection
        MyCars(CarName).KM = SavingDetails.CarsDetails(CarName).CarKM
        CarName = SavingDetails.CarsDetails(CarName).CarNumber
        CarImage.Left = SavingDetails.CarsDetails(CarName).CarLeft
        CarImage.Top = SavingDetails.CarsDetails(CarName).CarTop
    
        'Which picture to save for the Blue Car
        Select Case CarIntersection(CarName)
        Case 1, 2, 3, 7, 8, 9
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_left.bmp")
        Case 4, 5, 6, 10, 11, 12
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_right.bmp")
        Case 13, 14, 15, 19, 20, 21
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_down.bmp")
        Case 16, 17, 18, 22, 23, 24
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_up.bmp")
        Case 223
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_22.bmp")
        Case 1510
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_15.bmp")
        Case 921
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_9.bmp")
        Case 416
            CarImage.Picture = LoadPicture(App.Path & "\Images\" & CarPic & "_4.bmp")
        End Select
        
        CarWay(CarName) = "Move"
        CarTimer.Enabled = True
        
    End If
    
End Sub

Public Sub fnSavingDetails(CarName As Integer, CarImage As Image, CarSquare As Integer)      'Saving details

    SavingDetails.CarsDetails(CarName).CarFuel = MyCars(CarName).Fuel.FuelForNow
    SavingDetails.CarsDetails(CarName).CarIntersection = CarIntersection(CarName)
    SavingDetails.CarsDetails(CarName).CarKM = MyCars(CarName).KM
    SavingDetails.CarsDetails(CarName).CarNumber = CarName
    SavingDetails.CarsDetails(CarName).CarWay = CarWay(CarName)
    SavingDetails.CarsDetails(CarName).CarLeft = CarImage.Left
    SavingDetails.CarsDetails(CarName).CarTop = CarImage.Top
    SavingDetails.CarsDetails(CarName).CarSquare = CarSquare
    If ((CarImage.Top = CarsStartPosition(CarName).Top) And (CarImage.Left = CarsStartPosition(CarName).Left)) Then
        boolSaveInRest(CarName) = True
    Else: boolSaveInRest(CarName) = False
    End If

End Sub
