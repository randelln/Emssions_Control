' *********************************************************************************
'  Randell Naidoo
'  Class name: Diesel
' *********************************************************************************
Option Strict On
Option Infer Off
Option Explicit On

Public Class Eco
    Inherits Vehicles                 ' variables are declared
    Private _batterytype As String
    Private _batterysize As Double
    Private _besteco As String

    Public Property batterytype As String
        Get                                   ' property methods are created for the variables
            Return _batterytype
        End Get
        Set(value As String)
            _batterytype = value
        End Set
    End Property

    Public Property batterysize As Double
        Get
            Return _batterysize
        End Get
        Set(value As Double)
            _batterysize = value
        End Set
    End Property

    Public Property besteco As String
        Get
            Return _besteco
        End Get
        Set(value As String)
            _besteco = value
        End Set
    End Property

    Public Overrides Function display() As String
        Dim t As String = ""
        t &= model & vbNewLine & vbNewLine                     ' a function that returns a variable with the displaying of the information
        t &= "Weight: " & CStr(weight) & vbNewLine
        t &= "CO2 emmisions: " & CStr(0) & vbNewLine '0 because its Eco cars
        t &= "KM per Charge: " & CStr(KMpertank) & vbNewLine
        t &= "Battery Type: " & _batterytype & vbNewLine
        t &= "Battery size: " & CStr(_batterysize) & vbNewLine
        Return t
    End Function

    Public Overrides Function CalcFeulEmissions() As Double
        Return 0                    ' the calculation for the fuel emmissions of the electric powered cars
    End Function

End Class
