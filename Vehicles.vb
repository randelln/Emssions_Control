' *********************************************************************************
'  Randell Naidoo
'  Class name: Diesel
' *********************************************************************************
Option Strict On
Option Infer Off
Option Explicit On

Public Class Vehicles     ' variables
    Private _Model As String
    Private _weight As Double
    Private _CO2 As Double
    Private _Mileagepertank As Double

    Public Property KMpertank As Double    'property methods for the variables
        Get
            Return _Mileagepertank
        End Get
        Set(value As Double)
            _Mileagepertank = value
        End Set
    End Property

    Public Property model As String
        Get
            Return _Model
        End Get
        Set(value As String)
            _Model = value
        End Set
    End Property

    Public Property weight As Double
        Get
            Return _weight
        End Get
        Set(value As Double)
            _weight = value
        End Set
    End Property

    Public Property CO2 As Double
        Get
            Return _CO2
        End Get
        Set(value As Double)
            _CO2 = value
        End Set
    End Property


    Public Overridable Function CalcFeulEmissions() As Double        ' an overidable function for the base class that calculates the fuel emmisions of the vehicles
        Return CO2
    End Function

    Public Overridable Function Display() As String          ' a function that is overidable that displays the information
        Dim t As String = vbNewLine
        Return t
    End Function

End Class
