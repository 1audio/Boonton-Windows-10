Imports NationalInstruments.NI4882 'must be included to reference the LangInt assembly
Imports System.IO
Public Class KeithleyInterface

    'Private Shared TIMEOUT As Integer = c.T10s            'TIMEOUT assigned a value during initialization

    Private Const EOTMODE As Byte = 1                   ' Enable the END message
    Private Const EOSMODE As Byte = 0                   ' Disable the EOS mode
    Private Const ReadDelay As Integer = 1000           'Amount of time spent waiting after telling controller to detect something
    Private Const SNDelay As Integer = 2000      'Delay specific delay

    Private Const REFERENCEVOLTAGE As Double = 1 'Reference Voltage for converting V to and from DB

    Private thisBoonton As Device
    Private Boonton As Integer              'a reference to my instrument

    Private currentFrequency As Double = Double.NaN 'stored for setting the source

    Public Sub New(ByVal GPIBAddress As Integer, ByVal PrimaryAddress As Integer)
        thisBoonton = New Device(GPIBAddress, PrimaryAddress)

        Write("RE99", True)
    End Sub

    Public Sub New(ByVal GPIBAddress As Integer, ByVal PrimaryAddress As Integer, ByVal SecondaryAddress As Integer)
        thisBoonton = New Device(GPIBAddress, PrimaryAddress, SecondaryAddress)
        'Boonton = li.ibdev(GPIBAddress, PrimaryAddress, SecondaryAddress, TIMEOUT, EOTMODE, EOSMODE)
    End Sub

    'returns true if no errors
    Public Function Write(ByVal data As String, ByVal SeeError As Boolean) As Boolean
        If (Not data.Equals(String.Empty)) Then
            Try
                thisBoonton.Write(System.Text.Encoding.ASCII.GetBytes(data.ToCharArray()))
            Catch ex As Exception
                LastCommand = data
                ErrorObject = ex
                ErrorNumber = 1000
                Return False
            End Try
            If (SeeError) Then
                Dim Result As Integer = CheckError()
                'If there was a program error as a result of trying to check for the error,
                'the last command should stay as the check-error command (Since that's what caused it)
                If (Result = 1000) Then
                    Return False
                Else
                    LastCommand = data
                    Return Result = 0
                End If
            End If
        End If
        Return True
    End Function

    Private Function CheckError() As Integer
        Dim Success As Boolean = Write("TS", False)
        If (Success) Then
            ErrorNumber = CInt(Read().Trim("\"c, "r"c, "n"c))
        Else
            ErrorNumber = 1000
        End If

        Return ErrorNumber
    End Function

    Private LastCommand As String
    Private ErrorNumber As Integer
    Private ErrorObject As Exception = Nothing
    Public Function GetErrorNumber() As Integer
        Return ErrorNumber
    End Function

    Public Function GetErrorString() As String
        Select Case ErrorNumber
            Case 0
                Return "There is no error."
            Case 1
                Return LastCommand & " is an illegal source frequency entry."
            Case 2
                Return LastCommand & " is an illegal frequency step size entry."
            Case 3
                Return LastCommand & " is an illegal source level entry."
            Case 4
                Return LastCommand & " is an illegal level step size entry."
            Case 5
                Return LastCommand & " is an illegal special function entry."
            Case 6
                Return LastCommand & " is an illegal start frequency or level entry."
            Case 7
                Return LastCommand & " is an illegal stop frequency or level entry."
            Case 8
                Return LastCommand & " is an illegal low plot limit entry."
            Case 9
                Return LastCommand & " is an illegal high plot limit entry."
            Case 10
                Return LastCommand & " is an illegal bus address entry."
            Case 11
                Return LastCommand & " tried to recall an erased location or store in the read-only #99."
            Case 12
                Return LastCommand & " is an illegal voltage range or frequency entry."
            Case 13
                Return LastCommand & " is an illegal input voltage range."
            Case 14
                Return LastCommand & " is an illegal input range, notch frequency, or distortion range."
            Case 15
                Return LastCommand & " is an illegal input range, notch frequency, or SINAD range."
            Case 17
                Return LastCommand & " is an analyzer command, but the machine is in ratio mode."
            Case 18
                Return LastCommand & " invoked a ratio display overrange error."
            Case 22
                Return LastCommand & " is not a recognized command."
            'This one is made up for errors within the program
            Case 1000
                Return LastCommand & " resulted in the program error " & If(IsNothing(ErrorObject), "Unknown Error", ErrorObject.Message)
            Case Else
                Return LastCommand & " invoked an unknown error. Please look up the error code."
        End Select

    End Function
    'Returns Nothing if error
    Public Function Read() As String
        Dim result As String
        Try
            result = thisBoonton.ReadString()
            result = result.Replace(ControlChars.Lf, "\n").Replace(ControlChars.Cr, "\r")
        Catch ex As Exception
            ErrorObject = ex
            ErrorNumber = 1000
            Return Nothing
        End Try

        Return result
    End Function


    Public Function SetFloatingInput(ByVal SetTo As Boolean) As Boolean
        If (SetTo) Then
            Return Write("FA", True)
        Else
            Return Write("SA", True)
        End If
    End Function

    Public Function SetSourceDB(ByVal OutputVolts As Double) As Boolean
        Return Write("SL" & Str(OutputVolts) & "DB", True) 'source level to variable volts
    End Function

    Public Function SetSourceVolts(ByVal OutputVolts As Double) As Boolean
        Return Write("SL" & Str(OutputVolts) & "VO", True)
    End Function

    Public Function Tune(ByVal Frequency As Integer) As Boolean
        Return Write("DN" + CStr(Frequency) & "HZ", True)
    End Function

    Public Function SetSource(ByVal OutputVolts As Double, ByVal InverseRIAA As Boolean, ByVal val318uS As Boolean) As Boolean
        Dim dB3 As Double
        Dim RIAAOutput As Double

        CurrentOutputVolts = OutputVolts
        If (InverseRIAA And Not val318uS) Then 'no 3.18uS time constant
            If (Double.IsNaN(currentFrequency)) Then
                Throw New System.Exception("This implimentation expects Frequency to be set before Source.")
            End If
            dB3 = (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.000318 * 0.000318 * currentFrequency * currentFrequency))) - (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.00318 * 0.00318 * currentFrequency * currentFrequency))) - (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.000075 * 0.000075 * currentFrequency * currentFrequency)))
            dB3 = dB3 + 19.911018                   'normalizes to 1kHz
            RIAAOutput = Math.Pow(10, (dB3 / 20))   'converts from dB to voltage gain
            RIAAOutput = OutputVolts / RIAAOutput    'divides 1kHz input by voltage gain
            RIAAOutput = Math.Round(RIAAOutput, 5)  'rounds to tens of microvolts - the Boonton limit
            Return SetSourceDB(RIAAOutput) 'source level to variable volts
        ElseIf (InverseRIAA And val318uS) Then 'with Allen Wright 3.18uS time constant
            dB3 = (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.000318 * 0.000318 * currentFrequency * currentFrequency))) + (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.00000318 * 0.00000318 * currentFrequency * currentFrequency))) - (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.00318 * 0.00318 * currentFrequency * currentFrequency))) - (10 * Math.Log10(1 + (4 * Math.PI * Math.PI * 0.000075 * 0.000075 * currentFrequency * currentFrequency)))
            dB3 = dB3 + 19.909285
            RIAAOutput = Math.Pow(10, (dB3 / 20))
            RIAAOutput = OutputVolts / RIAAOutput
            RIAAOutput = Math.Round(RIAAOutput, 5)
            Return SetSourceDB(RIAAOutput) 'source level to variable volts
        End If

        Return SetSourceVolts(OutputVolts)
    End Function

    Public Function SetFrequency(ByVal Frequency As Integer) As Boolean
        currentFrequency = Frequency
        Dim test As Boolean = Write("SF" & Frequency & "HZ", True) ' set source frequency in Hz
        Return test
    End Function

    Public Function SetFilters(ByVal F As Integer, ByVal L As Integer) As Boolean
        Return Write("F" & CStr(F) & "L" & CStr(L), True)
    End Function

    Public Function MeasureVRMS() As Double
        If (Not Write("TVALVODB", False)) Then
            Return Double.NaN
        End If

        System.Threading.Thread.Sleep(ReadDelay)

        Dim Result As String = Read()
        If (IsNothing(Result)) Then
            Return Double.NaN
        End If

        If (CheckError() <> 0) Then
            Return Double.NaN
        End If

        'take left seven digits from the buffer, ie trim garbage to prevent math error
        Result = Strings.Left(Result, 7)
        Return System.Math.Round(Val(Result), 3)

    End Function

    Public Shared Function VoltsToDB(ByRef Volts As Double) As Double
        'This is the minimum setting for the boonton
        If (Volts < 0.000001) Then Return -120
        Return 20 * Math.Log10(Volts / REFERENCEVOLTAGE)
    End Function

    Public Shared Function DBToVolts(ByRef DB As Double) As Double
        Return Math.Pow(10, DB / 20) * REFERENCEVOLTAGE
    End Function

    'Brent added DB to set distortion as DB
    Public Function MeasureTHDN() As Double
        If (Not Write("TVDNPCDB", False)) Then
            Return Double.NaN
        End If

        System.Threading.Thread.Sleep(ReadDelay)

        Dim Result As String = Read()
        If (IsNothing(Result)) Then
            Return Double.NaN
        End If

        If (CheckError() <> 0) Then
            Return Double.NaN
        End If

        'take left seven digits from the buffer, ie trim garbage to prevent math error
        Result = Strings.Left(Result, 7)
        Return System.Math.Round(Val(Result), 3)
    End Function

    Public Function SetZero() As Boolean
        Write("RE99FA", True) ' source level zero volts
    End Function

    Public Function MeasureSN() As Double
        If (Not Write("SN", False)) Then
            Return Double.NaN
        End If
        System.Threading.Thread.Sleep((2000))

        Dim Result As String = Read()
        If (IsNothing(Result)) Then
            Return Double.NaN
        End If

        If (CheckError() <> 0) Then
            Return Double.NaN
        End If
        'take left seven digits from the buffer, ie trim garbage to prevent math error
        Result = Strings.Left(Result, 7)
        'rounds S/N to 1 significant digit
        Return System.Math.Round(Val(Result), 1)
    End Function

    Public Function Close() As Boolean
        Return SetZero()
    End Function

    'Based on function below, reverses dB number to roughly calculate watts
    Public Function dBToWatts(ByVal dB As Double, ByVal OhmLoad As Integer) As Double
        Dim InputVolts As Double = Math.Pow(10, dB / 20) * CurrentOutputVolts
        Return Math.Round(InputVolts * InputVolts / OhmLoad, 2)
    End Function

    ' Not yet implimented
    Private CurrentOutputVolts As Double
    Public Function MeasureVRMSConvertDB() As Boolean
        If (Not Write("ALV", False)) Then
            Return False
        End If
        Dim Result As String = Read()

        If (CheckError() <> 0) Then
            Return Double.NaN
        End If

        Result = Strings.Left(Result, 7)

        'rounds incoming input VRMS to the millivolt
        Dim InputVolts As Double = System.Math.Round(Val(Result), 3)
        Dim dB As Double

        'perform decibel math, ie 20log(V1/V2), then rounds to 2 significant digits
        dB = System.Math.Abs(InputVolts / CurrentOutputVolts)
        dB = System.Math.Log(dB) / System.Math.Log(10)
        dB = dB * 20
        dB = System.Math.Round(dB, 2)
    End Function
End Class
