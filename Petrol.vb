' *********************************************************************************
'  Randell Naidoo
'  Class name: Diesel
' *********************************************************************************
Option Strict On
Option Infer Off
Option Explicit On

Public Class Petrol
    Inherits Vehicles              ' variables are declared
    Private _fueltype As String
    Private _feultanksize As Double
    Private _bestpetrol As String
    Private _feulC As Double
    Private _EngineCap As Double
    Private _Cap As String

    Public Property cap As String
        Get                                      ' property methods are created for the variables
            Return _Cap
        End Get
        Set(value As String)
            _Cap = value
        End Set
    End Property

    Public Property fueltype As String
        Get
            Return _fueltype
        End Get
        Set(value As String)
            _fueltype = value
        End Set
    End Property

    Public Property fueltanksize As Double
        Get
            Return _feultanksize
        End Get
        Set(value As Double)
            _feultanksize = value
        End Set
    End Property

    Public Property bestpetrol As String
        Get
            Return _bestpetrol
        End Get
        Set(value As String)
            _bestpetrol = value
        End Set
    End Property

    Public Property feulcon As Double
        Get
            Return _feulC
        End Get
        Set(value As Double)
            _feulC = value
        End Set
    End Property

    Public Property enginecap As Double
        Get
            Return _EngineCap
        End Get
        Set(value As Double)
            _EngineCap = value
        End Set
    End Property

    Public Overrides Function display() As String         ' a function that returns a variable with the displaying of the information
        Dim t As String = ""
        t &= model & vbNewLine & vbNewLine
        t &= "Weight: " & CStr(weight) & "KG" & vbNewLine
        t &= "CO2 emmisions: " & CStr(CO2) & vbNewLine
        t &= "KM per tank: " & CStr(KMpertank) & "KM" & vbNewLine
        t &= "Fuel Type: " & _fueltype & vbNewLine
        t &= "Fuel Tank size: " & CStr(_feultanksize) & "ℓ" & vbNewLine
        t &= "Engine Capacity: " & cap & vbNewLine
        t &= "Fuel Consumption: " & Format(_feulC, "0.00") & " KM/ℓ" & vbNewLine
        Return t
    End Function

    Public Overrides Function CalcFeulEmissions() As Double
        Select Case enginecap             ' the calculation for the fuel emmissions of the petrol powered cars
            Case 0 To 1.4
                CO2 = 0.17
            Case 1.5 To 2.1
                CO2 = 0.22
            Case Else
                CO2 = 0.27
        End Select
        If fueltype = "95" Then
            Return CO2 - 0.03
        ElseIf fueltype = "LEADED" Then
            Return CO2 + 0.04
        Else
            Return CO2
        End If
    End Function

    Public Function FeulConsumption() As Double   'a function that calculates the fuel consumption of the petrol powered car
        _feulC = KMpertank / fueltanksize
        Return _feulC
    End Function
End Class
