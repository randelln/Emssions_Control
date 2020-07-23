' *********************************************************************************
'  Randell Naidoo
'  Class name: Diesel
' *********************************************************************************
Option Strict On
Option Infer Off
Option Explicit On

Public Class Diesel
    Inherits Vehicles
    Private _fueltype As String            ' variables are declared
    Private _feultanksize As Double
    Private _bestdiesel As String
    Private _feulC As Double
    Private _EngineCap As Double
    Private _Cap As String

    Public Property feulcon As Double           ' property methods are created for the variables
        Get
            Return _feulC
        End Get
        Set(value As Double)
            _feulC = value
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

    Public Property cap As String
        Get
            Return _Cap
        End Get
        Set(value As String)
            _Cap = value
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

    Public Property bestdiesel As String
        Get
            Return _bestdiesel
        End Get
        Set(value As String)
            _bestdiesel = value
        End Set
    End Property

    Public Property Engcap As Double
        Get
            Return _EngineCap
        End Get
        Set(value As Double)
            _EngineCap = value
        End Set
    End Property

    Public Overrides Function display() As String
        Dim t As String = ""                            ' a function that returns a variable with the displaying of the information
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
        Select Case Engcap                  ' the calculation for the fuel emmissions of the diesel powered cars
            Case 0 To 2.0
                CO2 = 0.12
            Case 2.1 To 4.0
                CO2 = 0.14
            Case Else
                CO2 = 0.18
        End Select
        If fueltype = "Low Sulphur Grade " Then
            Return CO2 - 0.04
        Else
            Return CO2
        End If
    End Function

    Public Function FeulConsumption() As Double
        _feulC = KMpertank / fueltanksize             'a function that calculates the fuel consumption of the diesel powered car
        Return _feulC
    End Function
End Class
