Attribute VB_Name = "mdlFuel"
Option Explicit
Option Base 1


Public Function Fueling2(CarName As Integer, CarImage As Image, CarTimer As Timer)
    'Setting the car back to te drive line
    frmIntersection.tmrFuel2.Enabled = True
    CarImage.Left = FuelStop2.Left
    CarImage.Top = Lines(1).TopEnd
    CarIntersection(CarName) = 1
    CarNeedFuel(CarName) = False
    boolFuel2Full = False
    
    Fueling2 = MyCars(CarName).Fuel.FullTank
    
End Function

Public Function Fueling1(CarName As Integer, CarImage As Image, CarTimer As Timer)
    'Setting the car back to te drive line
    frmIntersection.tmrFuel1.Enabled = True
    CarImage.Top = FuelStop1.Top
    CarImage.Left = Lines(18).LeftStart
    CarIntersection(CarName) = 18
    CarNeedFuel(CarName) = False
    boolFuel1Full = False
    
   Fueling1 = MyCars(CarName).Fuel.FullTank
    
End Function

Public Sub FuelLabel1(CarName As Integer)    'Shows the car who is fueling
    Select Case CarName
            Case 1
                frmIntersection.lblCarFuelName1.Caption = MyCars(BlueCar).CarName
            Case 2
                frmIntersection.lblCarFuelName1.Caption = MyCars(FireCar).CarName
            Case 3
                frmIntersection.lblCarFuelName1.Caption = MyCars(RedCar).CarName
            Case 4
                frmIntersection.lblCarFuelName1.Caption = MyCars(WhiteCar).CarName
            Case 5
                frmIntersection.lblCarFuelName1.Caption = MyCars(BlackCar).CarName
            Case 6
                frmIntersection.lblCarFuelName1.Caption = MyCars(GreenCar).CarName
            Case 7
                frmIntersection.lblCarFuelName1.Caption = MyCars(YellowCar).CarName
            Case 8
                frmIntersection.lblCarFuelName1.Caption = MyCars(TruckCar).CarName
    End Select
    
    frmIntersection.lblCarFuelName1.Visible = True
    frmIntersection.picSonol.Visible = True
    
End Sub

Public Sub FuelLabel2(CarName As Integer)    'Shows the car who is fueling
    
     Select Case CarName
            Case 1
                frmIntersection.lblCarFuelName2.Caption = MyCars(BlueCar).CarName
            Case 2
                frmIntersection.lblCarFuelName2.Caption = MyCars(FireCar).CarName
            Case 3
                frmIntersection.lblCarFuelName2.Caption = MyCars(RedCar).CarName
            Case 4
                frmIntersection.lblCarFuelName2.Caption = MyCars(WhiteCar).CarName
            Case 5
                frmIntersection.lblCarFuelName2.Caption = MyCars(BlackCar).CarName
            Case 6
                frmIntersection.lblCarFuelName2.Caption = MyCars(GreenCar).CarName
            Case 7
                frmIntersection.lblCarFuelName2.Caption = MyCars(YellowCar).CarName
            Case 8
                frmIntersection.lblCarFuelName2.Caption = MyCars(TruckCar).CarName
    End Select
    frmIntersection.lblCarFuelName2.Visible = True
    frmIntersection.picAlon.Visible = True
    
End Sub

Function CarClick(CarName As Integer, CarSquare As Integer, tmrCar As Timer) As Boolean  'Shows the car Details

    frmIntersection.lblCarFuel.Visible = True
    frmIntersection.lblCarKM.Visible = True
    frmIntersection.lblCarSpeed.Visible = True
    frmIntersection.tmrLableFade.Enabled = True
    frmIntersection.lblCarName.Caption = MyCars(CarName).CarName
    frmIntersection.lblCarFuel.Caption = "Fuel Left : " & MyCars(CarName).Fuel.FuelForNow
    frmIntersection.lblCarKM.Caption = "Km : " & MyCars(CarName).KM
    frmIntersection.lblCarSpeed.Caption = "Speed : " & CInt(1000 / tmrCar.Interval)
    
    CarClick = True
End Function
