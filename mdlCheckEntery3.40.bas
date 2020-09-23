Attribute VB_Name = "mdlCheckEntery"
Option Explicit
Option Base 1

'-------------------------------------------------------------------------------------------------------
'   This function set the car who start a new line in her position
'   checking thet there is no other car who started the same line in the same time
'-------------------------------------------------------------------------------------------------------


Function StartLinesCases(CarName As Integer, CarImage As Image, LineNumberBusy As Integer, LineNumber As Integer, LineSquare() As Boolean) As Boolean

    Dim CarCanGo As Boolean
    Dim i As Integer

    If LineBusy(LineNumberBusy) = False Then    '  If the line is not busy the car can start going
        LineBusy(LineNumberBusy) = True
        For i = 1 To intHaveToSlow
            LineSquare(i) = True
        Next i
        CarImage.Top = Lines(LineNumber).TopStart
        CarImage.Left = Lines(LineNumber).LeftStart
        CarWay(CarName) = "Move"
        'CarsInLine(LineNumberBusy) = CarsInLine(LineNumberBusy) + 1
        CarCanGo = True
    Else
        If ((LineSquare(2) = True) Or (LineSquare(intHaveToSlow)) = True) Then 'If line is busy
            CarWay(CarName) = "Start"                                          'Start this function again
            CarCanGo = False
        Else
            LineBusy(LineNumberBusy) = True
            For i = 1 To intHaveToSlow
                LineSquare(i) = True
            Next i
            CarImage.Top = Lines(LineNumber).TopStart
            CarImage.Left = Lines(LineNumber).LeftStart
            CarWay(CarName) = "Move"
            'CarsInLine(LineNumberBusy) = CarsInLine(LineNumberBusy) + 1
            CarCanGo = True
        End If
    End If
    
    StartLinesCases = CarCanGo
    
End Function

'Set the car in the line - after the traffic light

Function LinesCases(LineNumber As Integer, CarImage As Image) As String

    CarImage.Top = Lines(LineNumber).TopStart
    CarImage.Left = Lines(LineNumber).LeftStart
    
    LinesCases = "Move"

End Function

Public Sub MoveCars(Light1 As Integer, Light2 As Integer)

    Dim CarNumber As Integer
    
    For CarNumber = 1 To 8
    
        If ((CarIntersection(CarNumber) = Light1) Or (CarIntersection(CarNumber) = Light2)) Then
        
            CarWay(CarNumber) = "Move"
            Select Case CarNumber
                Case 1
                    frmIntersection.tmrBlueCar.Enabled = True
                Case 2
                    frmIntersection.tmrFireCar.Enabled = True
                Case 3
                    frmIntersection.tmrRedCar.Enabled = True
                Case 4
                    frmIntersection.tmrWhiteCar.Enabled = True
                Case 5
                    frmIntersection.tmrBlackCar.Enabled = True
                Case 6
                    frmIntersection.tmrGreenCar.Enabled = True
                Case 7
                    frmIntersection.tmrYellowCar.Enabled = True
                Case 8
                    frmIntersection.tmrTruckCar.Enabled = True
            End Select
        End If
        
    Next CarNumber
            
End Sub

