'' *********************************************************************************
'  Randell Naidoo
'  Class name: Diesel
' *********************************************************************************
Option Strict On
Option Infer Off
Option Explicit On

Public Class frmUN
    Private cars() As Vehicles
    Private Model, type, CapacityName As String        'variables are declared                     'Need all as the values change on runtime
    Private Weight, Mileage, Tank, Capacity, bestpc, bestpe, worstpc, worstpe, bestdc, bestde, worstdc, worstde, besteb, worsteb, bestee, worstee As Double
    Private counter As Integer = 0

    Private Sub txtPTank_TextChanged(sender As Object, e As EventArgs) Handles txtPTank.TextChanged      'Creating validaation and prescenece checks forcing the user to eneter the correct value before continuing
        cmbPLlitres.Visible = True
    End Sub

    Private Sub cmbPLlitres_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPLlitres.SelectedIndexChanged
        cmbPFuel.Visible = True
    End Sub

    Private Sub cmbPFuel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPFuel.SelectedIndexChanged
        btnPCreate.Visible = True
    End Sub

    Private Sub txtDTank_TextChanged(sender As Object, e As EventArgs) Handles txtDTank.TextChanged
        cmbDLitres.Visible = True
    End Sub

    Private Sub cmbDLitres_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDLitres.SelectedIndexChanged
        cmbDFuel.Visible = True
    End Sub

    Private Sub cmbDFuel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDFuel.SelectedIndexChanged
        btnDCreate.Visible = True
    End Sub

    Private Sub txtECapacity_TextChanged(sender As Object, e As EventArgs) Handles txtECapacity.TextChanged
        cmbEType.Visible = True
    End Sub

    Private Sub cmbEType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEType.SelectedIndexChanged
        btnECreate.Visible = True
    End Sub

    Public Sub Reset()
        rgpDiesel.Checked = False
        rgpPetrol.Checked = False
        rgpElectric.Checked = False
        txtModel.Clear()
        txtWeight.Clear()               'Reset all fields for a new entry
        txtMileage.Clear()
        cmbDFuel.SelectedIndex = -1
        cmbPFuel.SelectedIndex = -1
        cmbPLlitres.SelectedIndex = -1
        cmbDLitres.SelectedIndex = -1
        cmbEType.SelectedIndex = -1
        txtPTank.Clear()
        txtDTank.Clear()
        txtECapacity.Clear()
        btnGeneral.Visible = False

    End Sub

    Private Sub btnPCreate_Click(sender As Object, e As EventArgs) Handles btnPCreate.Click
        If ValidateNumber(txtPTank.Text) Then
            type = CStr(cmbPFuel.SelectedItem)
            Tank = CDbl(txtPTank.Text)
            Select Case cmbPLlitres.SelectedIndex
                Case 0
                    Capacity = 1.0
                    CapacityName = cmbPLlitres.SelectedItem.ToString
                Case 1
                    Capacity = 2.0                                        'Storing the Petrol information that is sent via PetrolCreate to create the object
                    CapacityName = cmbPLlitres.SelectedItem.ToString
                Case Else
                    Capacity = 3.0
                    CapacityName = cmbPLlitres.SelectedItem.ToString
            End Select
            PetrolCreate()
            btnPCreate.Visible = False
            cmbPFuel.Visible = False
            cmbPLlitres.Visible = False
            txtPTank.Visible = False
            rgpPetrol.Visible = True
            rgpDiesel.Visible = True
            rgpElectric.Visible = True
            txtModel.Visible = True
            txtMileage.Visible = True
            txtWeight.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label6.Visible = False
        Else
            MsgBox("Please enter a valid tank size.")
        End If
    End Sub

    Private Sub rgpPetrol_CheckedChanged(sender As Object, e As EventArgs) Handles rgpPetrol.CheckedChanged
        btnGeneral.Visible = True                                                                                            'Prescenece checks are forced here
    End Sub

    Private Sub rgpElectric_CheckedChanged(sender As Object, e As EventArgs) Handles rgpElectric.CheckedChanged
        btnGeneral.Visible = True
    End Sub

    Private Sub rgpDiesel_CheckedChanged(sender As Object, e As EventArgs) Handles rgpDiesel.CheckedChanged
        btnGeneral.Visible = True
    End Sub

    Private Sub btnDCreate_Click(sender As Object, e As EventArgs) Handles btnDCreate.Click
        If ValidateNumber(txtDTank.Text) Then
            type = CStr(cmbDFuel.SelectedItem)
            Tank = CDbl(txtDTank.Text)                       'Storing the Diesel information that is sent via PetrolCreate to create the object
            Select Case cmbDLitres.SelectedIndex
                Case 0
                    Capacity = 1.0
                    CapacityName = cmbDLitres.SelectedItem.ToString
                Case 1
                    Capacity = 3.0
                    CapacityName = cmbDLitres.SelectedItem.ToString
                Case Else
                    Capacity = 5.0
                    CapacityName = cmbDLitres.SelectedItem.ToString
            End Select
            DieselCreate()
            btnDCreate.Visible = False
            txtDTank.Visible = False
            cmbDLitres.Visible = False
            cmbDFuel.Visible = False
            rgpPetrol.Visible = True
            rgpDiesel.Visible = True
            rgpElectric.Visible = True
            txtModel.Visible = True
            txtMileage.Visible = True
            txtWeight.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label5.Visible = False
        Else
            MsgBox("Please enter a valid tank size.")
        End If
    End Sub

    Private Sub btnECreate_Click(sender As Object, e As EventArgs) Handles btnECreate.Click
        If ValidateNumber(txtECapacity.Text) Then
            type = CStr(cmbEType.SelectedItem)
            Capacity = CDbl(txtECapacity.Text)
            EcoCreate()
            btnECreate.Visible = False
            cmbEType.Visible = False
            txtECapacity.Visible = False
            rgpPetrol.Visible = True
            rgpDiesel.Visible = True                 'Storing the Electric Car information that is sent via PetrolCreate to create the object
            rgpElectric.Visible = True
            txtModel.Visible = True
            txtMileage.Visible = True
            txtWeight.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label7.Visible = False
        Else
            MsgBox("Please enter a valid tank size.")
        End If
    End Sub

    Public Sub PetrolCreate()      'sub-routine for the creation, inputs, calculations and display of a new petrol vehicle object
        counter = counter + 1
        ReDim Preserve cars(counter)
        Dim newpetrolcar As Petrol = New Petrol
        newpetrolcar.model = Model
        newpetrolcar.weight = Weight
        newpetrolcar.fueltype = type
        newpetrolcar.KMpertank = Mileage
        newpetrolcar.fueltanksize = Tank
        newpetrolcar.cap = CapacityName
        newpetrolcar.enginecap = Capacity
        newpetrolcar.CalcFeulEmissions()
        newpetrolcar.FeulConsumption()
        txtPetrol.Text &= newpetrolcar.display & vbNewLine
        If counter = 1 Then
            bestpc = newpetrolcar.CO2
            worstpc = newpetrolcar.CO2
        Else                                                  'Calcualting the best and worst fields and updating every time a new car is added
            If newpetrolcar.CO2 < bestpc Then
                bestpc = newpetrolcar.CO2
            Else
                If newpetrolcar.CO2 > worstpc Then
                    worstpc = newpetrolcar.CO2
                End If
            End If
        End If
        If counter = 1 Then
            bestpe = newpetrolcar.feulcon
            worstpe = newpetrolcar.feulcon
        Else
            If newpetrolcar.feulcon > bestpe Then
                bestpe = newpetrolcar.feulcon
            Else
                If newpetrolcar.feulcon < worstpe Then
                    worstpe = newpetrolcar.feulcon
                End If
            End If
        End If
        txtbpc.Text = CStr(bestpc)
        txtbpe.Text = CStr(bestpe)
        txtwpc.Text = CStr(worstpc)
        txtwpe.Text = CStr(worstpe)
        cars(counter) = newpetrolcar
        Reset()
        TabControl1.SelectedTab = tbsDisplay
    End Sub

    Private Sub frmUN_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Weight = 0
        Mileage = 0
        Tank = 0
        Capacity = 0
        bestpc = 0
        bestpe = 0
        worstpc = 0
        worstpe = 0          'Assigining values to 0 as they may be saved after the program closes
        bestdc = 0
        bestde = 0
        worstdc = 0
        worstde = 0
        besteb = 0
        worsteb = 0
        bestee = 0
        worstee = 0

    End Sub

    Public Sub DieselCreate()            'sub-routine for the creation, inputs, calculations and display of a new diesel vehicle object
        counter = counter + 1
        ReDim Preserve cars(counter)
        Dim newdiesel As Diesel = New Diesel
        newdiesel.model = Model
        newdiesel.weight = Weight
        newdiesel.fueltype = type
        newdiesel.KMpertank = Mileage
        newdiesel.fueltanksize = Tank
        newdiesel.Engcap = Capacity
        newdiesel.cap = CapacityName
        newdiesel.CalcFeulEmissions()
        newdiesel.FeulConsumption()
        txtDiesel.Text &= newdiesel.display & vbNewLine
        If counter = 1 Then
            bestdc = newdiesel.CO2
            worstdc = newdiesel.CO2
        Else
            If newdiesel.CO2 < bestdc Then
                bestdc = newdiesel.CO2
            Else
                If newdiesel.CO2 > worstdc Then
                    worstdc = newdiesel.CO2
                End If                                'Calcualting the best and worst fields and updating every time a new car is added
            End If
        End If
        If counter = 1 Then
            bestde = newdiesel.feulcon
            worstde = newdiesel.feulcon
        Else
            If newdiesel.feulcon > bestde Then
                bestde = newdiesel.feulcon
            Else
                If newdiesel.feulcon < worstde Then
                    worstde = newdiesel.feulcon
                End If
            End If
        End If
        txtbestdc.Text = CStr(bestdc)
        txtwde.Text = CStr(worstdc)
        txtbpc.Text = CStr(bestde)
        txtwpe.Text = CStr(worstde)
        cars(counter) = newdiesel
        Reset()
        TabControl1.SelectedTab = tbsDisplay
    End Sub

    Public Sub EcoCreate()               'sub-routine for the creation, inputs, calculations and display of a new electric vehicle object
        counter = counter + 1
        ReDim Preserve cars(counter)
        Dim neweco As Eco = New Eco
        neweco.model = Model
        neweco.weight = Weight
        neweco.batterytype = type
        neweco.KMpertank = Mileage
        neweco.batterysize = Capacity
        neweco.CalcFeulEmissions()
        txtEco.Text &= neweco.display & vbNewLine
        cars(counter) = neweco
        If counter = 1 Then
            besteb = neweco.batterysize
            worsteb = neweco.batterysize
        Else
            If neweco.batterysize > besteb Then
                besteb = neweco.batterysize
            Else
                If neweco.batterysize < worsteb Then
                    worsteb = neweco.batterysize
                End If
            End If
        End If
        txtbeb.Text = CStr(besteb)
        txtweb.Text = CStr(worsteb)
        Reset()
        TabControl1.SelectedTab = tbsDisplay
    End Sub

    Public Function ValidateNumber(ByRef num As String) As Boolean
        If IsNumeric(num) Then
            Return True
        Else                          'Function returning boolean to see if number is valid
            Return False
        End If
    End Function

    Private Sub btnGeneral_Click(sender As Object, e As EventArgs) Handles btnGeneral.Click
        Dim bNum As Boolean = False
        If (ValidateNumber(txtMileage.Text) = False) Then
            MsgBox("Please enter a valid Mileage")           'Validation 
            bNum = False
        Else
            If (ValidateNumber(txtWeight.Text) = False) Then
                MsgBox("Please enter a valid Weight")
                bNum = False
            Else
                bNum = True
            End If
        End If
        If bNum Then
            Model = txtModel.Text
            Mileage = CDbl(txtMileage.Text)
            Weight = CDbl(txtWeight.Text)
            rgpPetrol.Visible = False
            rgpDiesel.Visible = False
            rgpElectric.Visible = False
            txtModel.Visible = False
            txtMileage.Visible = False       'Ensuring precenence checks
            txtWeight.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            btnGeneral.Visible = False
            If rgpDiesel.Checked Then
                txtDTank.Visible = True
                Label5.Visible = True
            Else
                If rgpElectric.Checked Then
                    txtECapacity.Visible = True
                    Label7.Visible = True
                Else
                    txtPTank.Visible = True
                    Label6.Visible = True
                End If
            End If
        End If
    End Sub
End Class
