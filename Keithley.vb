Imports Microsoft.VisualBasic
Imports NationalInstruments.NI4882 'must be included to reference the LangInt assembly
Imports System.IO

Public MustInherit Class Keithley

    Private Const REFERENCEVOLTAGE As Double = 1 'Reference Voltage for converting V to and from DB

    'returns true if no errors
    Protected MustOverride Function Write(ByVal data As String) As Boolean
    'Returns Nothing if error
    Protected MustOverride Function Read() As String

    Public MustOverride Sub Close()




    Public Function SetDistortionType(ByVal Type As DistortionType) As Boolean
        Dim typeString As String
        Select Case Type
            Case DistortionType.THD
                typeString = "THD"
            Case DistortionType.THDN
                typeString = "THDN"
            Case DistortionType.SINAD
                typeString = "SINAD"
            Case Else
                Throw New System.Exception("I'm not quite sure how you got here...")
        End Select

        Return Write(":SENS:DIST:TYPE " + typeString)
    End Function

    Public Function SetAutoDetectFrequency(ByVal ShouldAutoDetect As Boolean) As Boolean
        Return Write(":SENS:DIST:FREQ:AUTO " & If(ShouldAutoDetect, "ON", "OFF"))
    End Function

    'This is done to measure distortion
    Public Function SetReadFrequency(ByVal Frequency As Integer) As Boolean
        Return Write(":SENS:DIST:FREQ " + CStr(Frequency))
    End Function

    'Like Autodetect, but it only needs to happen once per frequency so it doesn't slow the machine down.
    Public Function AquireFrequency() As Boolean
        Return Write(":SENS:DIST:FREQ:ACQ")
    End Function

    Public Function SetDistortionUnits(ByVal Units As Unit) As Boolean
        Dim UnitString As String
        Select Case Units
            Case Unit.Perc
                UnitString = "PERC"
            Case Unit.dB
                UnitString = "DB"
            Case Else
                Throw New System.Exception("Invalid unit")
        End Select

        Return Write(":UNIT:DIST " & UnitString)
    End Function

    Public Function SetFilter(ByVal TheFilter As Filter) As Boolean
        Dim FilterString As String
        Select Case TheFilter
            Case Filter.A
                FilterString = "A"
            Case Filter.CCIR
                FilterString = "CCIRARM"
            Case Filter.NONE
                FilterString = "NONE"
            Case Else
                Throw New System.Exception("I have no idea how you got here...")
        End Select

        Return Write(":SENS:DIST:SFIL " & FilterString)
    End Function

    Public Function PepareForDistortion() As Boolean
        Write("*RST")
        Write(":SENSE:FUNC 'DIST'")
        Write(":SENS:DIST:TYPE THDN")
        Write(":SENS:DIST:HARM 2")
        Write(":UNIT:DIST DB")
        SetFilter(Filter.NONE)
        Write(":SENS:DIST:RANGE:AUTO ON")
        Write(":OUTP:IMP HIZ")
        Write(":OUTP:CHAN2 ISINE")
        SetFrequency(1000)
        SetSource(1)
        Write(":OUTP ON")
        Return True
    End Function

    Public Function SetFrequency(ByVal frequency As Integer) As Boolean
        Return Write(":OUTP:FREQ " & CStr(frequency))
    End Function
    Public Function SetSource(ByVal OutputVolts As Double) As Boolean
        Return Write(":OUTP:AMPL " & Str(OutputVolts))
    End Function


    Public Function MeasureVRMS() As Double
        If (Not Write(":SENS:DIST:RMS?")) Then
            Return Double.NaN
        End If

        Dim result As String = Read()

        If (IsNothing(result)) Then
            Return Double.NaN
        End If

        'take left seven digits from the buffer, ie trim garbage to prevent math error
        result = result.Trim()
        Return VoltsToDB(CDbl(result))

    End Function

    Public Function MeasureTHDN() As Double
        If (Not Write("READ?")) Then
            Return Double.NaN
        End If

        Dim result As String = Read()

        If (IsNothing(result)) Then
            Return Double.NaN
        End If

        'take left seven digits from the buffer, ie trim garbage to prevent math error
        result = result.Trim()
        Return CDbl(result)

    End Function

    Public Enum Unit
        Perc
        dB
    End Enum

    Public Enum DistortionType
        THD
        THDN
        SINAD
    End Enum

    Public Enum Filter
        CCIR
        A
        NONE
    End Enum

    Public Shared Function VoltsToDB(ByRef Volts As Double) As Double
        'This is the minimum setting for the boonton
        If (Volts < 0.000001) Then Return -120
        Return 20 * Math.Log10(Volts / REFERENCEVOLTAGE)
    End Function

    Public Shared Function DBToVolts(ByRef DB As Double) As Double
        Return Math.Pow(10, DB / 20) * REFERENCEVOLTAGE
    End Function
End Class
' THD READING
' *RST ' Restore default settings.
' :SENS:FUNC 'DIST' ' Selects distortion function.
' :SENS:DIST:TYPE THD ' Selects THD distortion measurement type.
' :SENS:DIST:HARM 2 ' Sets the highest harmonic number to be ' measured. 
' :UNIT:DIST PERC ' Selects THD measurement units.
' :SENS:DIST:SFIL NONE ' Selects no shaping filter. 
' :SENS:DIST:RANG:AUTO ON ' Selects auto range for THD voltage input. 
' :OUTP:FREQ 1000 ' Sets the frequency of the source. 
' :OUTP:IMP HIZ ' Selects output impedance of source. 
' :OUTP:AMPL 1 ' Sets amplitude of source. 
' :OUTP:CHAN2 ISINE ' Selects inverted sine for second source. 
' :OUTP ON ' Turns output on. 
' :READ? ' Triggers one THD measurement and ' requests reading. 
' :SENS:DIST:RMS? ' Requests RMS volts reading for the TDH ' measurement. 
' :SENS:DIST:RMS?
' :SENS:DIST:THDN?

' Listing Fequencies to do
' 
'
' *RST ' Restores default conditions. *CLS ' Clears status registers. :STAT:OPER:ENAB 8 ' Causes the Operation Summary Bit to set ' when the sweep is done. Allows sweep end ' to be detected. *SRE 128 ' Enables Operation Summary Bit mask. ' Causes SRQ line to be asserted when the ' sweep is completed. :SENS:FUNC ‘DIST’ ' Selects distortion mode. :SENS:DIST:FREQ:AUTO OFF ' Turns off the AUTO frequency mode. :SENS:DIST:TYPE THDN ' Selects THD+noise mode. :SENS:DIST:LCO 500 ' Configures low cutoff to filter out ' noise below 500Hz. :SENS:DIST:LCO:STAT ON ' Turns on the low cutoff filter. :SENS:DIST:HCO 10000 ' Configures high cutoff to filter out ' noise above 10kHz. :SENS:DIST:HCO:STAT ON ' Turns on the high cutoff filter. :SENS:DIST:RANG 1 ' Sets the manual range for the sweep, ' since autoranging is not allowed. ' Change the range value as necessary. :OUTP:IMP HIZ ' Sets impedance of the output before ' creating the sweep list, so that the ' output has the correct amplitude. ' Change impedance as necessary. :OUTP:LIST 1,1000,1,1100,1,1200,1,1300,1,1400,1,1500,1,1600,           1,1700,1,1800,1,1900 ' Creates 10-point sweep list; 1kHz to ' 1.9kHz in 100Hz steps, one volt ' amplitude for all points. :OUTP:MODE LIST ' Selects sweep mode. :OUTP:LIST:DEL .1 ' Sets a source delay of 0.1 seconds. :OUTP:LIST:ELEM DIST,AMPL ' Selects distortion and amplitude as the ' data elements to be returned. :TRIG:COUN 10 ' Sets instrument to take 10 triggered ' measurements. :OUTP ON ' Turns the output on. :INIT ' Starts the sweep. ' (Wait for SRQ to be asserted - see NOTE at end of command sequence) :OUTPUT:LIST:DATA? ' Requests the sweep data. ' (Read data from 2015/2015-P - see NOTE that follows)


':SENS:DIST:HARM:MAGN? `start`,`end` ' Queries levels from starting ' harmonic to ending harmonic. ' `start` = 2 to 64 ' `end` = 2 to 64
