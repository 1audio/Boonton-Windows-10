Imports NationalInstruments.NI4882 'must be included to reference the LangInt assembly
Imports System.IO
Imports System.ComponentModel

Public Class KeithleyGPIB
    Inherits Keithley

    Private GPIBDevice As Device

    Public Sub New(ByVal GPIBAddress As Integer, ByVal PrimaryAddress As Integer)
        GPIBDevice = New Device(GPIBAddress, PrimaryAddress)
    End Sub

    Public Sub New(ByVal GPIBAddress As Integer, ByVal PrimaryAddress As Integer, ByVal SecondaryAddress As Integer)
        GPIBDevice = New Device(GPIBAddress, PrimaryAddress, SecondaryAddress)
        'Boonton = li.ibdev(GPIBAddress, PrimaryAddress, SecondaryAddress, TIMEOUT, EOTMODE, EOSMODE)
    End Sub


    'returns true if no errors
    Protected Overrides Function Write(ByVal data As String) As Boolean
        If (Not data.Equals(String.Empty)) Then
            Try
                GPIBDevice.Write(System.Text.Encoding.ASCII.GetBytes(data.ToCharArray()))
                System.Threading.Thread.Sleep(100)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly)
                Return False
            End Try
        End If
        Return True
    End Function


    'Returns Nothing if error
    Protected Overrides Function Read() As String
        Dim result As String
        Try
            result = GPIBDevice.ReadString()
            'result = result.Replace(ControlChars.Lf, "\n").Replace(ControlChars.Cr, "\r")
        Catch ex As Exception

            MsgBox(ex.Message, MsgBoxStyle.OkOnly)
            Return Nothing
        End Try

        Return result
    End Function


    'Closing the GPIB connection isn't nessisary
    Public Overrides Sub Close()
    End Sub

End Class