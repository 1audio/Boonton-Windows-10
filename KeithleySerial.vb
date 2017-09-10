Imports NationalInstruments.NI4882 'must be included to reference the LangInt assembly
Imports System.IO
Imports System.Text
Imports System.IO.Ports

Public Class KeithleySerial
    Inherits Keithley


    Private ComPort As IO.Ports.SerialPort = Nothing



    Public Sub New(ByVal COMName As String, ByVal Baud As Integer)
        Try
            ComPort = My.Computer.Ports.OpenSerialPort("COM1")
            ComPort.BaudRate = Baud
            ComPort.ReadTimeout = 10000
        Catch ex As TimeoutException
            If ComPort IsNot Nothing Then ComPort.Close()
            Throw ex
        End Try
    End Sub


    'returns true if no errors
    Protected Overrides Function Write(data As String) As Boolean
        If (Not data.Equals(String.Empty)) Then
            Try
                ComPort.Write(data & vbLf)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly)
                Return False
            End Try
        End If
        Return True
    End Function


    'Returns Nothing if error
    Protected Overrides Function Read() As String
        ' Receive strings from a serial port.
        Dim Result As String = Nothing

        Try
            'Give enough time for data to be sent
            System.Threading.Thread.Sleep(100)

            Dim Incoming As String = ComPort.ReadTo(vbCr)
            If Incoming IsNot Nothing Then
                Result = Incoming
            End If
        Catch ex As TimeoutException
            Throw ex
        End Try

        Return Result
    End Function

    Public Overrides Sub Close()
        ComPort.Close()
    End Sub
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
